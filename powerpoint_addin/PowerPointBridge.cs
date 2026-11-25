using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

[assembly: AssemblyTitle("ppt time PowerPoint Bridge")]
[assembly: AssemblyProduct("ppt time")]
[assembly: AssemblyVersion("1.0.2.0")]
[assembly: AssemblyFileVersion("1.0.2.0")]
[assembly: ComVisible(false)]

namespace PptTime.PowerPointBridge
{
    // The Office IDTExtensibility2 type library declares a dual interface and
    // SAFEARRAY(VARIANT)* for the final argument of each method.
    [ComVisible(true), Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IOfficeAddin
    {
        [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application,
            int connectMode, [In, MarshalAs(UnmanagedType.IDispatch)] object addin,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection(int removeMode,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
    }

    internal sealed class Page
    {
        public string Document = "";
        public long Window;
        public int? Current, Total, ShowPosition;
        public bool Ended;

        internal static Page Read(object window)
        {
            dynamic show = window;
            var page = new Page();
            try { page.Window = unchecked((uint)(int)show.HWND); } catch (Exception) { }
            dynamic presentation = null, view = null;
            try { presentation = show.Presentation; } catch (Exception) { }
            if (presentation != null)
            {
                try { page.Document = (string)presentation.FullName ?? ""; } catch (Exception) { }
                if (page.Document.Length == 0)
                    try { page.Document = (string)presentation.Name ?? ""; } catch (Exception) { }
                try { page.Total = (int)presentation.Slides.Count; } catch (Exception) { }
                if (page.Total <= 0) page.Total = null;
            }
            try { view = show.View; } catch (Exception) { }
            if (view != null)
            {
                try { page.Ended = (int)view.State == 5; } catch (Exception) { }
                // SlideIndex belongs to the actual presentation, even for a
                // custom show, skipped slides, transitions or a backward jump.
                // CurrentShowPosition belongs to the show sequence instead.
                if (!page.Ended)
                    try { page.Current = (int)view.Slide.SlideIndex; } catch (Exception) { }
            }
            if (page.Current < 1 || (page.Total.HasValue && page.Current > page.Total))
                page.Current = null;
            if (!page.Ended && !page.Current.HasValue && view != null)
            {
                try { page.ShowPosition = (int)view.CurrentShowPosition; } catch (Exception) { }
                if (page.ShowPosition < 1) page.ShowPosition = null;
            }
            return page;
        }
    }

    // Data is read on PowerPoint's UI apartment. Only these plain CLR
    // values are passed to the network worker, never an Office object.
    internal sealed class Report
    {
        public string session { get; set; }
        public long seq { get; set; }
        public int pid { get; set; }
        public bool active { get; set; }
        public bool ended { get; set; }
        public long hwnd { get; set; }
        public int? current { get; set; }
        public int? total { get; set; }
        public int? show_position { get; set; }
        public string document { get; set; }
        public string key { get; set; }
    }

    internal sealed class ShowTracker
    {
        private sealed class Live
        {
            internal readonly string Id = Guid.NewGuid().ToString("N");
            internal long Sequence;
            internal Page Last;
        }
        private readonly int pid;
        private readonly Action<Report> send;
        private readonly Dictionary<string, Live> shows = new Dictionary<string, Live>();
        private readonly Live idle = new Live { Last=new Page() };
        private readonly Stopwatch idleInterval = new Stopwatch();
        internal ShowTracker(int process, Action<Report> publish) { pid = process; send = publish; }

        private void Publish(Live entry, bool active)
        {
            Page page = entry.Last;
            send(new Report { session=entry.Id, seq=++entry.Sequence, pid=pid,
                active=active, ended=active && page.Ended, hwnd=page.Window,
                current=active ? page.Current : null, total=active ? page.Total : null,
                show_position=active ? page.ShowPosition : null,
                document=page.Document });
        }

        internal string Seen(object window)
        {
            Page page = Page.Read(window);
            string id;
            if (page.Window != 0)
                id = "window:" + page.Window.ToString(CultureInfo.InvariantCulture);
            else if (page.Document.Length != 0)
                id = "document:" + page.Document;
            else
                return null;
            Live entry;
            if (!shows.TryGetValue(id, out entry))
                shows[id] = entry = new Live();
            entry.Last = page;
            Publish(entry, true);
            return id;
        }

        internal void Sweep(ISet<string> present, bool connected=true)
        {
            foreach (string id in new List<string>(shows.Keys))
                if (!present.Contains(id))
                {
                    Publish(shows[id], false);
                    shows.Remove(id);
                }
            if (!connected || present.Count != 0)
            {
                idleInterval.Reset();
            }
            else if (!idleInterval.IsRunning || idleInterval.ElapsedMilliseconds >= 2000)
            {
                Publish(idle, false);
                idleInterval.Restart();
            }
        }
    }

    internal sealed class Publisher : IDisposable
    {
        private readonly object gate = new object();
        private readonly Dictionary<string, Report> pending = new Dictionary<string, Report>();
        private bool closing;
        private readonly Thread worker;
        private readonly string config;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();
        internal Publisher(string path)
        {
            config = path;
            worker = new Thread(Run) { IsBackground=true, Name="ppt time page reporting" };
            worker.Start();
        }

        internal void Send(Report value)
        {
            lock (gate)
            {
                if (closing) return;
                pending[value.session] = value;
                Monitor.Pulse(gate);
            }
        }
        public void Dispose()
        {
            lock (gate) { closing = true; Monitor.Pulse(gate); }
        }

        private void Run()
        {
            for (;;)
            {
                List<Report> batch;
                lock (gate)
                {
                    while (pending.Count == 0 && !closing) Monitor.Wait(gate);
                    if (closing && pending.Count == 0) return;
                    batch = new List<Report>(pending.Values);
                    pending.Clear();
                }
                // The endpoint and bearer rotate on app restart. Re-read once
                // per batch; the host sends an idle connectivity report every 2s.
                try
                {
                    if (!File.Exists(config)) continue;
                    string contents;
                    using (var f = new FileStream(config, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                    {
                        if (f.Length > 4096) continue;
                        using (var text = new StreamReader(f, Encoding.UTF8)) contents = text.ReadToEnd();
                    }
                    var entry = json.Deserialize<Dictionary<string,string>>(contents);
                    string token, raw;
                    Uri target;
                    if (entry == null || !entry.TryGetValue("key", out token) || token == null ||
                        token.Length < 32 || token.Length > 200 || !entry.TryGetValue("url", out raw) ||
                        !Uri.TryCreate(raw, UriKind.Absolute, out target) || target.Scheme != "http" ||
                        target.Host != "127.0.0.1" || target.AbsolutePath != "/report" ||
                        target.UserInfo != "" || target.Query != "" || target.Fragment != "") continue;
                    foreach (Report value in batch)
                    {
                        value.key = token;
                        byte[] body = Encoding.UTF8.GetBytes(json.Serialize(value));
                        var request = (HttpWebRequest)WebRequest.Create(target);
                        request.Proxy = null;
                        request.AllowAutoRedirect = false;
                        request.Method = "POST";
                        request.Timeout = 600;
                        request.ReadWriteTimeout = 600;
                        request.ContentType = "application/json";
                        request.ContentLength = body.Length;
                        using (var stream = request.GetRequestStream()) stream.Write(body, 0, body.Length);
                        try { using (request.GetResponse()) { } }
                        catch (WebException error)
                        {
                            if (error.Response != null) error.Response.Close();
                            break;
                        }
                    }
                }
                catch (Exception) { /* PowerPoint can run when ppt time is closed. */ }
            }
        }
    }

    [ComVisible(true), Guid("F7A81B92-8FBF-4535-B76E-62DD79855F90")]
    [ProgId("PptTime.PowerPointBridge"), ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IOfficeAddin))]
    public sealed class Connect : IOfficeAddin
    {
        // Verified against the PowerPoint 16 EApplication type information.
        private static readonly Guid Events = new Guid("914934C2-5A91-11CF-8700-00AA0060263B");
        private delegate void OfficeEvent([MarshalAs(UnmanagedType.IDispatch)] object item);
        private readonly Dictionary<int, OfficeEvent> subscriptions = new Dictionary<int, OfficeEvent>();
        private object app;
        private Publisher publisher;
        private ShowTracker tracker;
        private System.Windows.Forms.Timer timer;
        private bool reading;

        public void OnConnection(object application, int mode, object addin, ref Array custom)
        {
            try
            {
                Disconnect();
                app = application;
                string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                           "ppt time", "powerpoint-bridge.json");
                publisher = new Publisher(path);
                using (Process process = Process.GetCurrentProcess())
                    tracker = new ShowTracker(process.Id, publisher.Send);
                Subscribe(2011, OnSlide); // SlideShowBegin
                Subscribe(2013, OnSlide); // SlideShowNextSlide
                Subscribe(2014, OnEnd);   // SlideShowEnd
                timer = new System.Windows.Forms.Timer { Interval=500 };
                timer.Tick += OnTick;
                timer.Start();
                Refresh();
            }
            catch (Exception) { Disconnect(); }
        }

        private void Subscribe(int id, OfficeEvent handler)
        {
            try
            {
                ComEventsHelper.Combine(app, Events, id, handler);
                subscriptions[id] = handler;
            }
            catch (Exception) { /* The UI-thread heartbeat remains available. */ }
        }

        private void OnSlide(object window)
        {
            try { if (!reading && tracker != null) tracker.Seen(window); }
            catch (Exception) { }
        }
        private void OnEnd(object presentation) { Refresh(); }
        private void OnTick(object sender, EventArgs args) { Refresh(); }
        private void Refresh()
        {
            if (app == null || reading || tracker == null) return;
            reading = true;
            try
            {
                dynamic windows = ((dynamic)app).SlideShowWindows;
                int count = (int)windows.Count;
                var seen = new HashSet<string>();
                for (int i = 1; i <= count; i++)
                {
                    object window;
                    try { window = windows.Item(i); }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { window = windows[i]; }
                    string id = tracker.Seen(window);
                    if (id == null) return; // A transitioning window is still active.
                    seen.Add(id);
                }
                tracker.Sweep(seen);
            }
            catch (Exception) { /* Office rejects UI calls briefly around dialogs and transitions. */ }
            finally { reading = false; }
        }
        private void Disconnect()
        {
            if (timer != null) { timer.Stop(); timer.Tick -= OnTick; timer.Dispose(); timer = null; }
            foreach (var pair in subscriptions)
                try { ComEventsHelper.Remove(app, Events, pair.Key, pair.Value); } catch (Exception) { }
            subscriptions.Clear();
            if (tracker != null) { tracker.Sweep(new HashSet<string>(), false); tracker = null; }
            if (publisher != null) { publisher.Dispose(); publisher = null; }
            app = null;
        }
        public void OnDisconnection(int mode, ref Array custom) { Disconnect(); }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { Refresh(); }
        public void OnBeginShutdown(ref Array custom) { Disconnect(); }
    }
}
