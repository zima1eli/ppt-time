"""Check the EXE, MSI install plan and every payload byte without registering an app."""

import argparse
import ctypes
from ctypes import wintypes
import hashlib
import msilib
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import uuid

from win32com.client import Dispatch


def export_cabinet(package, path):
    """Copy the compressed stream using Windows Installer's byte-oriented API."""
    api = ctypes.WinDLL("msi")
    handle = wintypes.DWORD
    handles = []
    signatures = {
        "MsiOpenDatabaseW": [wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.POINTER(handle)],
        "MsiDatabaseOpenViewW": [handle, wintypes.LPCWSTR, ctypes.POINTER(handle)],
        "MsiViewExecute": [handle, handle],
        "MsiViewFetch": [handle, ctypes.POINTER(handle)],
        "MsiRecordReadStream": [handle, ctypes.c_uint, ctypes.c_void_p, ctypes.POINTER(wintypes.DWORD)],
        "MsiCloseHandle": [handle],
    }
    for name, signature in signatures.items():
        getattr(api, name).argtypes = signature
        getattr(api, name).restype = ctypes.c_uint

    def call(name, *args):
        code = getattr(api, name)(*args)
        if code:
            raise OSError(code, name)

    def open_handle(name, *args):
        value = handle()
        call(name, *args, ctypes.byref(value))
        handles.append(value)
        return value

    try:
        db = open_handle("MsiOpenDatabaseW", str(package), None)
        view = open_handle("MsiDatabaseOpenViewW", db,
                           "SELECT `Data` FROM `_Streams` WHERE `Name` = 'PptTimeFiles'")
        call("MsiViewExecute", view, 0)
        record = open_handle("MsiViewFetch", view)
        data = ctypes.create_string_buffer(1024 * 1024)
        with path.open("wb") as f:
            while True:
                size = wintypes.DWORD(len(data))
                call("MsiRecordReadStream", record, 1, data, ctypes.byref(size))
                if not size.value:
                    break
                f.write(data.raw[:size.value])
    finally:
        for value in reversed(handles):
            api.MsiCloseHandle(value)


def rows(db, query):
    view = db.OpenView(query)
    view.Execute(None)
    while (record := view.Fetch()) is not None:
        yield tuple(record.GetString(i + 1) for i in range(record.GetFieldCount()))
    view.Close()


def file_hash(path):
    with path.open("rb") as f:
        return hashlib.file_digest(f, "sha256").digest()


def verify_page_actions(db, payload, session):
    actions = {name: (int(kind), payload.get(source), target) for name, kind, source, target
               in rows(db, "SELECT `Action`, `Type`, `Source`, `Target` FROM `CustomAction`")}
    executable = Path("ppt time.exe")
    enable, remove, restore = "PptTimeEnablePages", "PptTimeRemovePages", "PptTimeRestorePages"
    wps_enable, wps_disable = "PptTimeEnableWpsPages", "PptTimeDisableWpsPages"
    wps_remove, wps_restore = "PptTimeRemoveWpsPages", "PptTimeRestoreWpsPages"
    assert actions[enable] == (1618, executable, "--setup-powerpoint-enable")
    assert actions[remove] == (1106, executable, "--setup-powerpoint-remove")
    assert actions[restore] == (1362, executable, "--setup-powerpoint-enable")
    assert actions[wps_enable] == (1618, executable, "--setup-wps-enable")
    assert actions[wps_disable] == (1618, executable, "--setup-wps-remove")
    assert actions[wps_remove] == (1106, executable, "--setup-wps-remove")
    assert actions[wps_restore] == (1362, executable, "--setup-wps-enable")
    assert ("FindPptTimeWpsOwner", "1", r"Software\ppt time\WpsPresentationBridge", "RegistrationOwner", "2") in tuple(
        rows(db, "SELECT `Signature_`, `Root`, `Key`, `Name`, `Type` FROM `RegLocator`"))
    assert ("PPTTIME_PRIOR_WPS_OWNER", "FindPptTimeWpsOwner") in tuple(rows(db, "SELECT `Property`, `Signature_` FROM `AppSearch`"))
    plan = {name: (int(number), condition) for name, condition, number in rows(db,
            "SELECT `Action`, `Condition`, `Sequence` FROM `InstallExecuteSequence`")}
    for action in (enable, wps_enable, wps_disable):
        assert plan["InstallFiles"][0] < plan[action][0] < plan["InstallFinalize"][0]
    assert plan[restore][0] < plan[remove][0] < plan["RemoveFiles"][0]
    assert plan[wps_restore][0] < plan[wps_remove][0] < plan["RemoveFiles"][0]
    properties = ("Installed", "REMOVE", "OLDER_PPT_TIME", "UPGRADINGPRODUCTCODE", "PPTTIME_ENABLE_POWERPOINT",
                  "PPTTIME_PRIOR_OWNER", "PPTTIME_PRIOR_LOAD", "PPTTIME_ENABLE_WPS", "PPTTIME_PRIOR_WPS_OWNER")
    page_actions = (enable, remove, restore, wps_enable, wps_disable, wps_remove, wps_restore)
    cases = [
        ({}, set()),                                                       # quiet install
        ({"PPTTIME_ENABLE_WPS": "0"}, set()),                               # fresh, unchecked
        ({"PPTTIME_ENABLE_POWERPOINT": "1"}, {enable}),
        ({"PPTTIME_ENABLE_WPS": "1"}, {wps_enable}),
        ({"OLDER_PPT_TIME": "{old}", "PPTTIME_ENABLE_POWERPOINT": "1", "PPTTIME_ENABLE_WPS": "1"}, {enable, wps_enable}),
        ({"Installed": "1", "REMOVE": "ALL", "UPGRADINGPRODUCTCODE": "{new}", "PPTTIME_PRIOR_WPS_OWNER": "setup"}, set()),
        ({"OLDER_PPT_TIME": "{old}", "PPTTIME_PRIOR_WPS_OWNER": "setup"}, set()),  # no new WPS choice
        ({"OLDER_PPT_TIME": "{old}", "PPTTIME_PRIOR_WPS_OWNER": "setup", "PPTTIME_ENABLE_WPS": "0"}, {wps_disable}),
        ({"OLDER_PPT_TIME": "{old}", "PPTTIME_PRIOR_WPS_OWNER": "manual", "PPTTIME_ENABLE_WPS": "0"}, set()),
        ({"Installed": "1", "REMOVE": "ALL", "PPTTIME_PRIOR_WPS_OWNER": "setup"}, {remove, wps_remove, wps_restore}),
        ({"Installed": "1", "REMOVE": "ALL", "PPTTIME_PRIOR_WPS_OWNER": "disabled"}, {remove}),
        ({"Installed": "1", "REMOVE": "ALL", "PPTTIME_PRIOR_OWNER": "manual",
          "PPTTIME_PRIOR_LOAD": "#3"}, {remove}),
        ({"Installed": "1", "REMOVE": "ALL", "PPTTIME_PRIOR_OWNER": "setup",
          "PPTTIME_PRIOR_LOAD": "#2"}, {remove}),
        ({"Installed": "1", "REMOVE": "ALL", "PPTTIME_PRIOR_OWNER": "setup",
          "PPTTIME_PRIOR_LOAD": "#3"}, {remove, restore}),
    ]
    property_id = session._oleobj_.GetIDsOfNames("Property")
    for values, expected in cases:
        for name in properties:
            session._oleobj_.Invoke(property_id, 0, 4, 0, name, values.get(name, ""))
        results = {name: session.EvaluateCondition(plan[name][1]) for name in page_actions}
        assert all(result in (0, 1) for result in results.values()), (values, results)
        actual = {name for name, result in results.items() if result == 1}
        assert actual == expected, (values, actual, expected)


def verify(stage, exe):
    stage, exe = stage.resolve(strict=True), exe.resolve(strict=True)
    temp = Path(tempfile.gettempdir()) / ("ppt-time-verify-" + uuid.uuid4().hex)
    temp.mkdir()
    try:
        msi, cab, files = temp / "app.msi", temp / "payload.cab", temp / "extracted"
        subprocess.run([str(exe), "/extract", str(msi)], check=True)
        db = msilib.OpenDatabase(str(msi), msilib.MSIDBOPEN_READONLY)
        properties = dict(rows(db, "SELECT `Property`, `Value` FROM `Property`"))
        assert "ALLUSERS" not in properties
        assert properties["ProductName"] == "ppt time"
        assert {"PPTTIME_ENABLE_WPS", "PPTTIME_PRIOR_WPS_OWNER"} <= set(properties["SecureCustomProperties"].split(";"))
        assert db.GetSummaryInformation(0).GetProperty(msilib.PID_WORDCOUNT) & 8
        directory = {key: (parent, longname.split("|")[-1]) for key, parent, longname in
                     rows(db, "SELECT `Directory`, `Directory_Parent`, `DefaultDir` FROM `Directory`")}
        assert directory["INSTALLDIR"] == ("UserPrograms", "ppt time")
        assert directory["UserPrograms"] == ("LocalAppDataFolder", "Programs")
        component_dirs = dict(rows(db, "SELECT `Component`, `Directory_` FROM `Component`"))
        components = dict(rows(db, "SELECT `Component`, `KeyPath` FROM `Component`"))
        payload = {}
        for file_id, component, filename in rows(db, "SELECT `File`, `Component_`, `FileName` FROM `File`"):
            assert components[component] == file_id
            rel = Path(filename.split("|")[-1])
            folder = component_dirs[component]
            while folder != "INSTALLDIR":
                folder, name = directory[folder]
                rel = Path(name) / rel
            payload[file_id] = rel
        assert set(payload.values()) == {p.relative_to(stage) for p in stage.rglob("*") if p.is_file()}
        sequences = dict(rows(db, "SELECT `Action`, `Sequence` FROM `InstallExecuteSequence`"))
        assert int(sequences["InstallInitialize"]) < int(sequences["RemoveExistingProducts"]) < int(sequences["InstallFiles"])
        assert all(name in sequences for name in ("RemoveFiles", "RemoveShortcuts", "RemoveRegistryValues", "UnpublishFeatures"))
        shortcuts = list(rows(db, "SELECT `Directory_`, `Target` FROM `Shortcut`"))
        assert len(shortcuts) == 1 and shortcuts[0][0] == "ProgramMenuFolder"
        assert payload[shortcuts[0][1][2:-1]].suffix.lower() == ".exe"
        assert list(rows(db, "SELECT `Root` FROM `Registry`")) == [("1",)]
        assert len(list(rows(db, "SELECT `ActionProperty` FROM `Upgrade`"))) == 2
        # Calculate actual destination directories with the operating system's installer.
        installer = Dispatch("WindowsInstaller.Installer")
        session = installer.OpenPackage(str(msi), 1)  # ignore currently installed product state
        for action in ("CostInitialize", "FileCost", "CostFinalize"):
            if session.DoAction(action) != 1:
                raise RuntimeError("Windows Installer could not cost the package: " + action)
        accessor = session._oleobj_
        install_dir = accessor.Invoke(accessor.GetIDsOfNames("TargetPath"), 0, 2, 1, "INSTALLDIR")
        assert Path(install_dir).parts[-2:] == ("Programs", "ppt time"), install_dir
        verify_page_actions(db, payload, session)
        db.Close()
        del accessor, session, installer
        export_cabinet(msi, cab)
        files.mkdir()
        expand = Path(os.environ["WINDIR"]) / "System32" / "expand.exe"
        subprocess.run([str(expand), str(cab), "-F:*", str(files)], check=True, stdout=subprocess.DEVNULL)
        for file_id, rel in payload.items():
            assert file_hash(stage / rel) == file_hash(files / file_id), rel
        print(f"Verified {len(payload)} files, per-user install path {install_dir}, Start menu shortcut, upgrade and uninstall tables.")
    finally:
        shutil.rmtree(temp)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path)
    parser.add_argument("exe", type=Path)
    args = parser.parse_args()
    verify(args.stage, args.exe)
