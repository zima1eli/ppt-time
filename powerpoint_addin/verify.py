"""Validate the compiled bridge in x86 and x64 CLR using a synthetic Office model.

Uses an ephemeral, credential-checked local HTTP endpoint and no Office installation
or registry writes. Run on Windows with CPython and a .NET Framework 4.x compiler.
"""
import json
import os
from pathlib import Path
import secrets
import shutil
import subprocess
import sys
import tempfile
import threading
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from build import compiler, NAME, ROOT


HARNESS = r'''
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

public class Slide { public int SlideIndex { get; set; } }
public class Slides { public int Count { get; set; } }
public class Doc { public string FullName { get; set; } public Slides Slides { get; set; } }
public class View {
    public int CurrentShowPosition {get {return 2;}}
    public Slide Slide {get;set;}
    public int State {get;set;}
}
public class Show {
    public int HWND {get;set;} public View View {get;set;} public Doc Presentation {get;set;}
}
public class WindowCollection {
    public List<Show> Items = new List<Show>();
    public bool Ready;
    public int Count {get{if(!Ready)throw new InvalidOperationException();return Items.Count;}}
    public Show Item(int oneBased){return Items[oneBased-1];}
}
public class OfficeApp {public WindowCollection SlideShowWindows {get;set;}}

public class NativeProbe {
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    public delegate int StartupComplete(IntPtr self, ref IntPtr array);
    static BindingFlags All = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;
    static Type Type(Assembly a,string name){return a.GetType("PptTime.PowerPointBridge."+name,true);}
    static void Require(bool v,string msg){if(!v)throw new Exception(msg);}
    static void Invoke(object o,string method,params object[] args){
        o.GetType().GetMethod(method,All).Invoke(o,args);
    }
    static Show Make(int h,int position,int count){
        return new Show {HWND=h, View=new View {State=1,Slide=new Slide{SlideIndex=position}},
                         Presentation=new Doc{FullName="fixture-"+h+".pptx",Slides=new Slides{Count=count}}};
    }
    static void Yield(){Thread.Sleep(700);}
    public static int Main(string[] args){
        try {
            Assembly a=Assembly.LoadFrom(args[0]);
            Type t=Type(a,"Connect"), iface=Type(a,"IOfficeAddin");
            object bridge=Activator.CreateInstance(t);
            Require(Marshal.GenerateProgIdForType(t)=="PptTime.PowerPointBridge","ProgID");
            Require(t.GUID==new Guid("F7A81B92-8FBF-4535-B76E-62DD79855F90"),"class GUID");
            Require(iface.GUID==new Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"),"host interface GUID");
            PortableExecutableKinds pe;ImageFileMachine machine;a.ManifestModule.GetPEKind(out pe,out machine);
            Require((pe & PortableExecutableKinds.ILOnly)!=0 &&
                    (pe & PortableExecutableKinds.Required32Bit)==0 &&
                    (pe & PortableExecutableKinds.PE32Plus)==0,"DLL must run inside x86 and x64 PowerPoint");
            IntPtr unk=Marshal.GetIUnknownForObject(bridge);IntPtr dev;Guid iid=iface.GUID;
            Require(Marshal.QueryInterface(unk,ref iid,out dev)==0,"COM interface not exposed");
            try {
                IntPtr slot=Marshal.ReadIntPtr(Marshal.ReadIntPtr(dev),10*IntPtr.Size);
                var startup=(StartupComplete)Marshal.GetDelegateForFunctionPointer(slot,typeof(StartupComplete));
                IntPtr empty=IntPtr.Zero;
                Require(startup(dev,ref empty)==0,"native vtable callback failed");
            } finally{Marshal.Release(dev);Marshal.Release(unk);}
            Array custom=new object[0];var windows=new WindowCollection();
            var app=new OfficeApp{SlideShowWindows=windows};
            Invoke(bridge,"OnConnection",app,1,null,custom);
            object old=t.GetField("publisher",All).GetValue(bridge);
            Require(old!=null,"OnConnection failed");
            // Replace default discovery location only for this fixture; no user profile writes.
            ((IDisposable)old).Dispose();
            Type pub=Type(a,"Publisher");
            object publisher=Activator.CreateInstance(pub,All,null,new object[]{args[1]},null);
            t.GetField("publisher",All).SetValue(bridge,publisher);
            Type report=Type(a,"Report");
            Delegate cb=Delegate.CreateDelegate(typeof(Action<>).MakeGenericType(report),
                                                 publisher,pub.GetMethod("Send",All));
            object tracker=Activator.CreateInstance(Type(a,"ShowTracker"),All,null,new object[]{12345,cb},null);
            t.GetField("tracker",All).SetValue(bridge,tracker);
            windows.Ready=true;
            Invoke(bridge,"OnTick",null,EventArgs.Empty); Yield();
            var show=Make(101,1,53);windows.Items.Add(show);
            Invoke(bridge,"OnSlide",show); Yield();
            show.View.Slide.SlideIndex=16;
            Invoke(bridge,"OnSlide",show); Yield();
            var other=Make(202,4,6);windows.Items.Add(other);
            // No slide event for this change. Pump the application's STA timer.
            show.View.Slide.SlideIndex=3;
            for(int i=0;i<18;i++){System.Windows.Forms.Application.DoEvents();Thread.Sleep(50);}
            var savedSlides=show.Presentation.Slides;
            show.Presentation.Slides=null;
            show.View.Slide.SlideIndex=7;
            Invoke(bridge,"OnSlide",show); Yield();
            show.Presentation.Slides=savedSlides;
            show.View.Slide=null;
            Invoke(bridge,"OnSlide",show); Yield();
            show.View.State=5;
            Invoke(bridge,"OnTick",null,EventArgs.Empty); Yield();
            windows.Items.Remove(show);
            Invoke(bridge,"OnEnd",show.Presentation); Yield();
            Invoke(bridge,"OnDisconnection",1,custom); Yield();
            Console.WriteLine("CLR "+(8*IntPtr.Size)+"-bit, native vtable and synthetic page reports passed");
            return 0;
        }catch(Exception e){Console.Error.WriteLine(e);return 1;}
    }
}
'''


def main():
    dll = Path(sys.argv[1] if len(sys.argv) > 1 else ROOT / (NAME + ".dll")).resolve(strict=True)
    directory = Path(tempfile.gettempdir()) / ("ppt-time-addin-check-" + uuid.uuid4().hex)
    directory.mkdir()
    received = []
    key = secrets.token_urlsafe(32)

    class Receiver(BaseHTTPRequestHandler):
        def do_POST(self):
            size = int(self.headers.get("Content-Length", 0))
            data = json.loads(self.rfile.read(size))
            valid = self.path == "/report" and data.pop("key", None) == key
            if valid:
                received.append(data)
            self.send_response(204 if valid else 403)
            self.end_headers()

        def log_message(self, *args):
            pass

    server = ThreadingHTTPServer(("127.0.0.1", 0), Receiver)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        endpoint = directory / "test.json"
        endpoint.write_text(json.dumps({"key": key, "url": f"http://127.0.0.1:{server.server_port}/report"}), encoding="utf-8")
        source = directory / "fixture.cs"
        source.write_text(HARNESS, encoding="ascii")
        architectures = ["x86"]
        if os.environ.get("PROCESSOR_ARCHITEW6432") or sys.maxsize > 2**32:
            architectures.append("x64")
        for arch in architectures:
            received.clear()
            exe = directory / f"test-{arch}.exe"
            subprocess.run([str(compiler()), "/nologo", "/platform:" + arch, "/target:exe",
                            "/r:System.Windows.Forms.dll", "/out:" + str(exe), str(source)], check=True)
            subprocess.run([str(exe), str(dll), str(endpoint)], check=True, timeout=25)
            first = [s for s in received if s["hwnd"] == 101]
            other = [s for s in received if s["hwnd"] == 202]
            assert any(s["hwnd"] == 0 and not s["active"] for s in received), "connected idle report"
            assert other and first, "both separate show windows should be reported"
            assert {s["current"] for s in first if s["active"] and not s["ended"]} == {1, 16, 3, 7, None}, [
                (s["seq"], s["current"], s["total"], s["active"], s["ended"]) for s in first]
            assert all(s["total"] == 53 for s in first if s["active"] and s["current"] != 7)
            assert any(s["current"] == 7 and s["total"] is None and s["show_position"] is None for s in first)
            assert any(s["current"] is None and s["show_position"] == 2 and s["total"] == 53 for s in first)
            assert all(s["show_position"] is None for s in first if s["ended"] or not s["active"])
            assert any(s["ended"] and s["active"] and s["current"] is None for s in first)
            assert any(s["current"] == 4 and s["total"] == 6 for s in other)
            assert not first[-1]["active"] and not other[-1]["active"]
            for group in (first, other):
                assert len({s["session"] for s in group}) == 1
                assert [s["seq"] for s in group] == sorted({s["seq"] for s in group})
            assert first[0]["session"] != other[0]["session"]
            print(f"{arch} loopback captured real index with/without total, independent custom-show position, two windows, end and close.")
    finally:
        server.shutdown()
        server.server_close()
        thread.join()
        shutil.rmtree(directory)


if __name__ == "__main__":
    main()
