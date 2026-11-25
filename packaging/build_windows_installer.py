"""Build a per-user ppt time EXE installer using Windows CPython 3.11 and .NET Framework."""

import argparse
import hashlib
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import uuid

import msilib
from msilib import schema, sequence


PRODUCT = "ppt time"
FAMILY = uuid.UUID("6beed204-88c7-4e62-a6b1-e8ea674074dd")


def guid(key):
    return "{" + str(uuid.uuid5(FAMILY, key)).upper() + "}"


def identifier(prefix, path):
    return prefix + hashlib.sha256(path.as_posix().casefold().encode()).hexdigest()[:32]


def build_msi(stage, msi, main, version, icon):
    db = msilib.init_database(str(msi), schema, PRODUCT, guid("version:" + version), version, PRODUCT)
    info = db.GetSummaryInformation(20)
    info.SetProperty(msilib.PID_WORDCOUNT, 10)  # compressed long names; standard user privileges
    info.SetProperty(msilib.PID_PAGECOUNT, 500)
    info.Persist()
    msilib.add_data(db, "Property", [
        ("UpgradeCode", "{" + str(FAMILY).upper() + "}"),
        ("ARPNOMODIFY", "1"),
        ("SecureCustomProperties",
         "OLDER_PPT_TIME;NEWER_PPT_TIME;PPTTIME_ENABLE_POWERPOINT;PPTTIME_PRIOR_OWNER;PPTTIME_PRIOR_LOAD;"
         "PPTTIME_ENABLE_WPS;PPTTIME_PRIOR_WPS_OWNER"),
    ])
    msilib.add_data(db, "Upgrade", [
        ("{" + str(FAMILY).upper() + "}", None, version, None, 0, None, "OLDER_PPT_TIME"),
        ("{" + str(FAMILY).upper() + "}", version, None, None, 2, None, "NEWER_PPT_TIME"),
    ])
    msilib.add_data(db, "LaunchCondition", [
        ("Installed OR NOT NEWER_PPT_TIME", "A newer release of ppt time is already installed."),
    ])
    # Upgrade inside the install transaction so that failure rolls back.
    positions = {"FindRelatedProducts": 25, "RemoveExistingProducts": 1510}
    for table in ("AdminExecuteSequence", "AdminUISequence", "AdvtExecuteSequence",
                  "InstallExecuteSequence", "InstallUISequence"):
        rows = [(name, condition, positions.get(name, pos))
                for name, condition, pos in getattr(sequence, table)]
        msilib.add_data(db, table, rows)
    msilib.add_data(db, "Directory", [
        ("TARGETDIR", None, "SourceDir"), ("LocalAppDataFolder", "TARGETDIR", "."),
        ("UserPrograms", "LocalAppDataFolder", "Programs"),
        ("INSTALLDIR", "UserPrograms", "PPTTIM~1|ppt time"),
        ("ProgramMenuFolder", "TARGETDIR", "."),
    ])
    msilib.add_data(db, "Feature", [("Complete", None, PRODUCT, None, 1, 1, "INSTALLDIR", 0)])
    cab = msilib.CAB("PptTimeFiles")
    dirs = {Path("."): "INSTALLDIR"}
    primary_file = None
    entries = sorted((path.relative_to(stage) for path in stage.rglob("*")),
                     key=lambda path: (len(path.parts), path.as_posix().casefold()))
    for number, rel in enumerate(entries):
        full = stage / rel
        if full.is_symlink() or (full.stat().st_file_attributes & 0x400):
            raise ValueError(f"Unsupported reparse point in staging: {rel}")
        # Supply explicit unique DOS aliases so long and case-preserving names survive.
        parent = dirs[rel.parent]
        shortname = f"T{number:07X}"
        longname = f"{shortname}|{rel.name}"
        if full.is_dir():
            directory = identifier("D", rel)
            msilib.add_data(db, "Directory", [(directory, parent, longname)])
            dirs[rel] = directory
            continue
        file_id = identifier("F", rel)
        component = identifier("C", rel)
        idx, _ = cab.append(str(full), rel.name, file_id)
        msilib.add_data(db, "Component", [(component, guid("file:" + rel.as_posix().casefold()),
                                           parent, 256, None, file_id)])
        msilib.add_data(db, "FeatureComponents", [("Complete", component)])
        msilib.add_data(db, "File", [(file_id, component, longname, full.stat().st_size,
                                      None, None, 512, idx)])
        if rel == main:
            primary_file = file_id
    if not primary_file:
        raise ValueError(f"Staged executable is missing: {main}")

    # File-type actions run with the installing user's identity. Registration is
    # optional; a blocked Office add-in must not prevent installing the timer.
    # Commit enable only after all application files have installed successfully.
    file_exe, deferred, commit, rollback, continue_on_error = 18, 1024, 512, 256, 64
    msilib.add_data(db, "CustomAction", [
        ("PptTimeEnablePages", file_exe | deferred | commit | continue_on_error,
         primary_file, "--setup-powerpoint-enable"),
        ("PptTimeRestorePages", file_exe | deferred | rollback | continue_on_error,
         primary_file, "--setup-powerpoint-enable"),
        ("PptTimeRemovePages", file_exe | deferred | continue_on_error,
         primary_file, "--setup-powerpoint-remove"),
        ("PptTimeEnableWpsPages", file_exe | deferred | commit | continue_on_error,
         primary_file, "--setup-wps-enable"),
        ("PptTimeDisableWpsPages", file_exe | deferred | commit | continue_on_error,
         primary_file, "--setup-wps-remove"),
        ("PptTimeRestoreWpsPages", file_exe | deferred | rollback | continue_on_error,
         primary_file, "--setup-wps-enable"),
        ("PptTimeRemoveWpsPages", file_exe | deferred | continue_on_error,
         primary_file, "--setup-wps-remove"),
    ])
    user = 1
    # The HKCU preference and Office Addins keys are shared between registry views.
    msilib.add_data(db, "RegLocator", [
        ("FindPptTimeBridgeOwner", user, r"Software\ppt time\PowerPointBridge", "RegistrationOwner", 2),
        ("FindPptTimeBridgeLoad", user, r"Software\Microsoft\Office\PowerPoint\Addins\PptTime.PowerPointBridge",
         "LoadBehavior", 2),
        ("FindPptTimeWpsOwner", user, r"Software\ppt time\WpsPresentationBridge", "RegistrationOwner", 2),
    ])
    msilib.add_data(db, "AppSearch", [
        ("PPTTIME_PRIOR_OWNER", "FindPptTimeBridgeOwner"), ("PPTTIME_PRIOR_LOAD", "FindPptTimeBridgeLoad"),
        ("PPTTIME_PRIOR_WPS_OWNER", "FindPptTimeWpsOwner"),
    ])
    removing = 'Installed AND REMOVE~="ALL" AND NOT UPGRADINGPRODUCTCODE'
    msilib.add_data(db, "InstallExecuteSequence", [
        ("PptTimeRestorePages", removing + ' AND PPTTIME_PRIOR_OWNER="setup" AND PPTTIME_PRIOR_LOAD="#3"', 3480),
        ("PptTimeRemovePages", removing, 3490),
        ("PptTimeEnablePages", 'PPTTIME_ENABLE_POWERPOINT="1" AND NOT REMOVE', 6500),
        ("PptTimeRestoreWpsPages", removing + ' AND PPTTIME_PRIOR_WPS_OWNER="setup"', 3481),
        ("PptTimeRemoveWpsPages", removing + ' AND PPTTIME_PRIOR_WPS_OWNER="setup"', 3491),
        ("PptTimeEnableWpsPages", 'PPTTIME_ENABLE_WPS="1" AND NOT REMOVE', 6501),
        ("PptTimeDisableWpsPages", 'PPTTIME_ENABLE_WPS="0" AND PPTTIME_PRIOR_WPS_OWNER="setup" AND NOT REMOVE', 6502),
    ])

    # A per-user shortcut uses an HKCU registry key path for repair and uninstall.
    msilib.add_data(db, "Component", [("MenuComponent", guid("start-menu"), "ProgramMenuFolder",
                                       260, None, "MenuKey")])
    msilib.add_data(db, "FeatureComponents", [("Complete", "MenuComponent")])
    msilib.add_data(db, "Registry", [("MenuKey", 1, r"Software\ppt time\Installer", "StartMenu",
                                      "#1", "MenuComponent")])
    icon_ref = None
    if icon:
        icon_ref = "ppttime.ico"
        msilib.add_data(db, "Icon", [(icon_ref, msilib.Binary(str(icon)))])
        msilib.add_data(db, "Property", [("ARPPRODUCTICON", icon_ref)])
    msilib.add_data(db, "Shortcut", [("StartMenuApp", "ProgramMenuFolder", "PPTTIM~1|ppt time",
                                      "MenuComponent", f"[#{primary_file}]", None, PRODUCT,
                                      None, icon_ref, 0 if icon else None, 1, "INSTALLDIR")])
    cab.commit(db)
    db.Commit()
    db.Close()


def build(stage, output, executable, version, icon=None):
    stage, output = Path(stage).resolve(strict=True), Path(output).resolve()
    main = Path(executable)
    if main.name != str(main) or main.suffix.lower() != ".exe" or not (stage / main).is_file():
        raise ValueError("Specify the executable name in the root of the staged app.")
    if output.suffix.lower() != ".exe" or output.is_relative_to(stage):
        raise ValueError("Write the Setup.exe outside the application staging directory.")
    if not re.fullmatch(r"\d+\.\d+\.\d+", version):
        raise ValueError("Windows Installer version requires three numbers, e.g. 1.0.0.")
    if any(int(n) > limit for n, limit in zip(version.split("."), (255, 255, 65535))):
        raise ValueError("Version exceeds Windows Installer limits.")
    if not msilib.AMD64:
        raise RuntimeError("Use Windows x64 CPython 3.11.")
    compiler = Path(os.environ["WINDIR"]) / "Microsoft.NET" / "Framework64" / "v4.0.30319" / "csc.exe"
    if not compiler.is_file():
        raise FileNotFoundError(compiler)
    output.parent.mkdir(parents=True, exist_ok=True)
    # Use an ordinary inherited ACL in a private uniquely named build directory.
    build_id = uuid.uuid4().hex
    scratch = Path(tempfile.gettempdir()) / ("ppt-time-installer-" + build_id)
    publication = output.with_name(f"{output.name}.{build_id}.building")
    scratch.mkdir()
    try:
        msi = scratch / "ppt-time.msi"
        build_msi(stage, msi, main, version, icon)
        bootstrap = scratch / "Setup.exe"
        metadata = scratch / "Version.cs"
        metadata.write_text("using System.Reflection;\n" +
                            f'[assembly: AssemblyVersion("{version}.0")]\n' +
                            f'[assembly: AssemblyFileVersion("{version}.0")]\n', encoding="ascii")
        arguments = [str(compiler), "/nologo", "/target:winexe", "/platform:x64", "/optimize+",
                     "/codepage:65001", "/reference:System.Windows.Forms.dll",
                     f"/resource:{msi},ppt-time.msi", f"/out:{bootstrap}"]
        if icon:
            arguments.append(f"/win32icon:{icon}")
        arguments.append(str(metadata))
        arguments.append(str(Path(__file__).with_name("Setup.cs")))
        subprocess.run(arguments, check=True)
        # Publish only after both the MSI and bootstrap build succeeded.
        shutil.copyfile(bootstrap, publication)
        os.replace(publication, output)
    finally:
        shutil.rmtree(scratch)
        publication.unlink(missing_ok=True)
    return output


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--executable", default="ppt time.exe")
    parser.add_argument("--version", required=True)
    parser.add_argument("--icon", type=Path)
    options = parser.parse_args()
    print(build(options.stage, options.output, options.executable, options.version, options.icon))
