"""Build the optional PowerPoint bridge using the Windows .NET Framework compiler."""
import os
from pathlib import Path
import subprocess
import sys
import uuid


ROOT = Path(__file__).resolve().parent
NAME = "PptTime.PowerPointBridge"


def compiler():
    framework = Path(os.environ["WINDIR"]) / "Microsoft.NET"
    for architecture in ("Framework64", "Framework"):
        csc = framework / architecture / "v4.0.30319" / "csc.exe"
        if csc.is_file():
            return csc
    raise FileNotFoundError("Install .NET Framework 4.x to build the PowerPoint add-in.")


def build(output_dir):
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    # CSC derives the assembly identity from the output basename.
    scratch = output_dir / (".compile-" + uuid.uuid4().hex)
    scratch.mkdir()
    binary = scratch / (NAME + ".dll")
    try:
        subprocess.run([str(compiler()), "/nologo", "/target:library", "/platform:anycpu",
                        "/optimize+", "/codepage:65001", "/r:Microsoft.CSharp.dll",
                        "/r:System.Windows.Forms.dll", "/r:System.Web.Extensions.dll",
                        "/out:" + str(binary), str(ROOT / "PowerPointBridge.cs")], check=True)
        target = output_dir / binary.name
        os.replace(binary, target)
        return target
    finally:
        binary.unlink(missing_ok=True)
        scratch.rmdir()


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else ROOT))
