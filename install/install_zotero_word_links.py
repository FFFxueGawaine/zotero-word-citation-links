from __future__ import annotations

import shutil
import sys
import tempfile
import time
import zipfile
import os
from pathlib import Path
from xml.etree import ElementTree as ET
import winreg

try:
    import pythoncom
    import win32com.client
except ImportError as exc:  # pragma: no cover - startup guidance
    print("INSTALL_FAILED: 缺少依赖 pywin32。", file=sys.stderr)
    print("请先运行: python -m pip install pywin32", file=sys.stderr)
    raise


NS = {"ui": "http://schemas.microsoft.com/office/2006/01/customui"}
ET.register_namespace("", NS["ui"])

SCRIPT_DIR = Path(__file__).resolve().parent
APPDATA_DIR = Path(os.environ.get("APPDATA", Path.home()))
DEFAULT_TEMPLATE = APPDATA_DIR / "Microsoft" / "Word" / "STARTUP" / "Zotero.dotm"
DEFAULT_BAS = SCRIPT_DIR / "ZoteroWordHyperlinks.bas"
BACKUP_DIR = SCRIPT_DIR / "backup"
BACKUP_NAME = "Zotero.backup.before-linking.dotm"
SECURITY_KEY = r"Software\Microsoft\Office\16.0\Word\Security"

SEPARATOR_ID = "ZoteroCitationLinksSeparator"
CREATE_ID = "ZoteroCreateCitationLinksButton"
REMOVE_ID = "ZoteroRemoveCitationLinksButton"
UNLINK_ID = "ZoteroRemoveCodes"
REFRESH_ID = "RefreshZotero"


def read_access_vbom_state() -> tuple[bool, int]:
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, SECURITY_KEY) as key:
        try:
            value, _ = winreg.QueryValueEx(key, "AccessVBOM")
            return True, int(value)
        except FileNotFoundError:
            return False, 0


def set_access_vbom(value: int) -> None:
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, SECURITY_KEY) as key:
        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, int(value))


def restore_access_vbom(existed: bool, value: int) -> None:
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, SECURITY_KEY) as key:
        if existed:
            winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, int(value))
        else:
            try:
                winreg.DeleteValue(key, "AccessVBOM")
            except FileNotFoundError:
                pass


def get_active_word():
    try:
        return win32com.client.GetActiveObject("Word.Application")
    except Exception:
        return None


def find_addin(app, template_path: Path):
    if app is None:
        return None
    target = str(template_path).lower()
    for i in range(1, app.AddIns.Count + 1):
        addin = app.AddIns(i)
        current = str(Path(addin.Path) / addin.Name).lower()
        if current == target:
            return addin
    return None


def wait_until_unlocked(path: Path, timeout: float = 15.0) -> None:
    deadline = time.time() + timeout
    last_error = None
    while time.time() < deadline:
        try:
            with path.open("rb"):
                return
        except OSError as exc:
            last_error = exc
            time.sleep(0.2)
    raise RuntimeError(f"Template is still locked: {last_error}")


def backup_template(template_path: Path) -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    backup_path = BACKUP_DIR / BACKUP_NAME
    shutil.copy2(template_path, backup_path)
    return backup_path


def patch_custom_ui(template_path: Path) -> None:
    temp_fd, temp_name = tempfile.mkstemp(suffix=".dotm", prefix="zotero_link_patch_", dir=str(template_path.parent))
    os.close(temp_fd)
    Path(temp_name).unlink(missing_ok=True)
    temp_path = Path(temp_name)

    with zipfile.ZipFile(template_path, "r") as src_zip:
        custom_ui_bytes = src_zip.read("customUI/customUI.xml")
        root = ET.fromstring(custom_ui_bytes.decode("utf-8"))
        group = root.find(".//ui:group[@id='ZoteroGroup']", NS)
        if group is None:
            raise RuntimeError("ZoteroGroup was not found in customUI.xml")

        for child in list(group):
            child_id = child.attrib.get("id")
            if child_id in {SEPARATOR_ID, CREATE_ID, REMOVE_ID}:
                group.remove(child)

        refresh_button = None
        unlink_button = None
        for child in list(group):
            child_id = child.attrib.get("id")
            if child_id == REFRESH_ID:
                refresh_button = child
            elif child_id == UNLINK_ID:
                unlink_button = child

        if refresh_button is None:
            raise RuntimeError("RefreshZotero button was not found in customUI.xml")
        if unlink_button is None:
            raise RuntimeError("ZoteroRemoveCodes button was not found in customUI.xml")

        refresh_button.set("onAction", "ZoteroWordHyperlinks.ZoteroRefreshAndCreateCitationLinks")
        refresh_button.set(
            "supertip",
            "Update all citations to reflect changes, then rebuild citation links",
        )

        group.remove(unlink_button)
        children = list(group)
        refresh_index = children.index(refresh_button)
        group.insert(refresh_index + 1, unlink_button)

        separator = ET.Element(f"{{{NS['ui']}}}separator", {"id": SEPARATOR_ID})
        create_button = ET.Element(
            f"{{{NS['ui']}}}button",
            {
                "id": CREATE_ID,
                "label": "Create Citation Links",
                "imageMso": "HyperlinkInsert",
                "onAction": "ZoteroWordHyperlinks.ZoteroCreateCitationLinks",
                "supertip": "Create clickable links from Zotero citations to bibliography entries",
                "keytip": "K",
            },
        )
        remove_button = ET.Element(
            f"{{{NS['ui']}}}button",
            {
                "id": REMOVE_ID,
                "label": "Remove Citation Links",
                "imageMso": "TableUnlinkExternalData",
                "onAction": "ZoteroWordHyperlinks.ZoteroRemoveCitationLinks",
                "supertip": "Remove citation links and bibliography bookmarks created by the hyperlink helper",
                "keytip": "L",
            },
        )

        children = list(group)
        unlink_index = children.index(unlink_button)
        group.insert(unlink_index + 1, separator)
        group.insert(unlink_index + 2, create_button)
        group.insert(unlink_index + 3, remove_button)

        updated_custom_ui = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(temp_path, "w") as dst_zip:
            for info in src_zip.infolist():
                data = src_zip.read(info.filename)
                if info.filename == "customUI/customUI.xml":
                    data = updated_custom_ui
                dst_zip.writestr(info, data)

    shutil.move(temp_path, template_path)


def import_macro_module(template_path: Path, bas_path: Path) -> None:
    pythoncom.CoInitialize()
    app = None
    doc = None
    try:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(str(template_path), ReadOnly=False, AddToRecentFiles=False, Visible=False)
        project = doc.VBProject
        components = project.VBComponents

        for index in range(components.Count, 0, -1):
            component = components.Item(index)
            if component.Name == "ZoteroWordHyperlinks":
                components.Remove(component)

        components.Import(str(bas_path))
        doc.Save()
    finally:
        if doc is not None:
            doc.Close(False)
        if app is not None:
            app.Quit(False)
        pythoncom.CoUninitialize()


def install_into_running_word(app, template_path: Path) -> None:
    if app is None:
        return

    try:
        addin = find_addin(app, template_path)
    except Exception:
        app = get_active_word()
        addin = find_addin(app, template_path)
        if app is None:
            return

    if addin is None:
        app.AddIns.Add(FileName=str(template_path), Install=True)
        return

    addin.Installed = True


def ensure_addin_installed(template_path: Path, timeout: float = 10.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        app = get_active_word()
        if app is None:
            time.sleep(0.5)
            continue
        try:
            install_into_running_word(app, template_path)
            addin = find_addin(app, template_path)
            if addin is not None and addin.Installed:
                return True
        except Exception:
            pass
        time.sleep(0.5)
    return False


def uninstall_from_running_word(app, template_path: Path) -> bool:
    if app is None:
        return False
    addin = find_addin(app, template_path)
    if addin is None or not addin.Installed:
        return False
    addin.Installed = False
    return True


def main() -> int:
    template_path = DEFAULT_TEMPLATE
    bas_path = DEFAULT_BAS

    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")
    if not bas_path.exists():
        raise FileNotFoundError(f"BAS module not found: {bas_path}")

    original_access_vbom = read_access_vbom_state()
    active_word = get_active_word()

    was_unloaded = False
    try:
        set_access_vbom(1)

        was_unloaded = uninstall_from_running_word(active_word, template_path)
        if was_unloaded:
            wait_until_unlocked(template_path)

        backup_path = backup_template(template_path)
        patch_custom_ui(template_path)
        import_macro_module(template_path, bas_path)

        if active_word is not None and not ensure_addin_installed(template_path):
            raise RuntimeError("Template was updated, but Zotero.dotm could not be reloaded into the running Word session")

        print(f"Backup created: {backup_path}")
        print(f"Template updated: {template_path}")
        print("Ribbon buttons installed: Create Citation Links, Remove Citation Links")
        return 0
    finally:
        if active_word is not None and was_unloaded:
            try:
                ensure_addin_installed(template_path)
            except Exception:
                pass
        restore_access_vbom(*original_access_vbom)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"INSTALL_FAILED: {exc}", file=sys.stderr)
        raise
