from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import textwrap
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
import winreg

try:
    import pythoncom
    import win32com.client
except ImportError as exc:  # pragma: no cover - startup guidance
    print("BUILD_FAILED: missing pywin32.", file=sys.stderr)
    print("Please run: python -m pip install pywin32", file=sys.stderr)
    raise


NS = {"ui": "http://schemas.microsoft.com/office/2006/01/customui"}
ET.register_namespace("", NS["ui"])

REPO_ROOT = Path(__file__).resolve().parent.parent
INSTALL_DIR = REPO_ROOT / "install"
MAC_DIR = REPO_ROOT / "mac"
DOCS_DIR = REPO_ROOT / "docs"
DIST_DIR = REPO_ROOT / "dist"

UPSTREAM_REPO_URL = "https://github.com/zotero/zotero-word-for-mac-integration.git"
UPSTREAM_COMMIT = "084f16d78b4924b15ad24c83dd5d8aaaeb9998d8"
UPSTREAM_TEMPLATE_RELATIVE = Path("install") / "Zotero.dotm"
UPSTREAM_COPYING_RELATIVE = Path("COPYING")

SEPARATOR_ID = "ZoteroCitationLinksSeparator"
CREATE_ID = "ZoteroCreateCitationLinksButton"
REMOVE_ID = "ZoteroRemoveCitationLinksButton"
SET_COLOR_ID = "ZoteroSetLinkColorButton"
UNLINK_ID = "ZoteroRemoveCodes"
REFRESH_ID = "RefreshZotero"
SECURITY_KEY = r"Software\Microsoft\Office\16.0\Word\Security"

OUTPUT_ZIP = DIST_DIR / "zotero-word-links-mac-template.zip"
PACKAGE_ROOT = "zotero-word-links-mac-template"
MAC_EXECUTABLE_SUFFIXES = {".command"}


def run_git(args: list[str], cwd: Path | None = None) -> str:
    result = subprocess.run(
        ["git", *args],
        cwd=str(cwd) if cwd else None,
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    return result.stdout.strip()


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


def patch_custom_ui(template_path: Path) -> None:
    temp_fd, temp_name = tempfile.mkstemp(
        suffix=".dotm",
        prefix="zotero_mac_patch_",
        dir=str(template_path.parent),
    )
    os.close(temp_fd)
    Path(temp_name).unlink(missing_ok=True)
    temp_path = Path(temp_name)

    with zipfile.ZipFile(template_path, "r") as src_zip:
        custom_ui_bytes = src_zip.read("customUI/customUI.xml")
        root = ET.fromstring(custom_ui_bytes.decode("utf-8"))
        group = root.find(".//ui:group[@id='ZoteroGroup']", NS)
        if group is None:
            raise RuntimeError("ZoteroGroup was not found in upstream Mac customUI.xml")

        for child in list(group):
            child_id = child.attrib.get("id")
            if child_id in {SEPARATOR_ID, CREATE_ID, REMOVE_ID, SET_COLOR_ID}:
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
            raise RuntimeError("RefreshZotero button was not found in upstream Mac customUI.xml")
        if unlink_button is None:
            raise RuntimeError("ZoteroRemoveCodes button was not found in upstream Mac customUI.xml")

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


def clone_upstream_template(temp_root: Path) -> tuple[Path, str]:
    clone_dir = temp_root / "upstream"
    run_git(["clone", UPSTREAM_REPO_URL, str(clone_dir)])
    run_git(["checkout", UPSTREAM_COMMIT], cwd=clone_dir)
    commit = run_git(["rev-parse", "HEAD"], cwd=clone_dir)
    return clone_dir, commit


def write_upstream_info(package_dir: Path, upstream_commit: str) -> None:
    info_path = package_dir / "UPSTREAM_TEMPLATE_INFO.md"
    info_text = textwrap.dedent(
        f"""\
        # Upstream Template Info

        This package contains a modified `Zotero.dotm` derived from Zotero's official Mac Word template.

        - Upstream repository: {UPSTREAM_REPO_URL}
        - Upstream commit: `{upstream_commit}`
        - Upstream template path: `{UPSTREAM_TEMPLATE_RELATIVE.as_posix()}`
        - Upstream license: AGPLv3 (see `UPSTREAM_COPYING.txt`)

        The hyperlink helper additions in this project are integrated into the upstream template so that Word for Mac users can manually install the template.
        """
    )
    info_path.write_text(info_text, encoding="utf-8")


def write_restore_note(package_dir: Path) -> None:
    restore_path = package_dir / "RESTORE_MAC.txt"
    restore_text = textwrap.dedent(
        """\
        Mac Restore Steps

        1. Quit Microsoft Word.
        2. Delete the replaced Zotero.dotm from the Word Startup template folder.
        3. Copy your original backup Zotero.dotm back into the same folder.
        4. Reopen Word.
        """
    )
    restore_path.write_text(restore_text, encoding="utf-8")


def add_file_to_zip(archive: zipfile.ZipFile, source_path: Path, arcname: str) -> None:
    info = zipfile.ZipInfo.from_file(source_path, arcname)
    if source_path.suffix in MAC_EXECUTABLE_SUFFIXES:
        info.create_system = 3
        info.external_attr = 0o100755 << 16
    else:
        info.create_system = 3
        info.external_attr = 0o100644 << 16

    with source_path.open("rb") as fh:
        archive.writestr(info, fh.read(), compress_type=zipfile.ZIP_DEFLATED)


def verify_custom_ui(template_path: Path) -> None:
    with zipfile.ZipFile(template_path, "r") as archive:
        xml_text = archive.read("customUI/customUI.xml").decode("utf-8")
    if CREATE_ID not in xml_text or REMOVE_ID not in xml_text:
        raise RuntimeError("Mac template build verification failed: customUI buttons not found")
    if SET_COLOR_ID in xml_text:
        raise RuntimeError("Mac template build verification failed: retired Set Link Color button is still present")


def verify_macro_module(template_path: Path) -> None:
    pythoncom.CoInitialize()
    app = None
    doc = None
    try:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(str(template_path), ReadOnly=True, AddToRecentFiles=False, Visible=False)
        project = doc.VBProject
        names = [project.VBComponents.Item(i).Name for i in range(1, project.VBComponents.Count + 1)]
        if "ZoteroWordHyperlinks" not in names:
            raise RuntimeError("Mac template build verification failed: VBA module not found")
    finally:
        if doc is not None:
            doc.Close(False)
        if app is not None:
            app.Quit(False)
        pythoncom.CoUninitialize()


def build_package() -> Path:
    DIST_DIR.mkdir(parents=True, exist_ok=True)
    bas_path = INSTALL_DIR / "ZoteroWordHyperlinks.bas"
    mac_install_doc = MAC_DIR / "MAC_INSTALL.md"
    mac_install_script = MAC_DIR / "install_mac.command"
    mac_restore_script = MAC_DIR / "restore_mac.command"
    style_guide = DOCS_DIR / "STYLE_GUIDE.md"
    if not bas_path.exists():
        raise FileNotFoundError(f"Macro module not found: {bas_path}")
    if not mac_install_doc.exists():
        raise FileNotFoundError(f"Mac install guide not found: {mac_install_doc}")
    if not mac_install_script.exists():
        raise FileNotFoundError(f"Mac install script not found: {mac_install_script}")
    if not mac_restore_script.exists():
        raise FileNotFoundError(f"Mac restore script not found: {mac_restore_script}")
    if not style_guide.exists():
        raise FileNotFoundError(f"Style guide not found: {style_guide}")

    temp_root = Path(tempfile.mkdtemp(prefix="zotero_word_links_mac_build_"))
    access_vbom_existed, access_vbom_value = read_access_vbom_state()
    try:
        upstream_dir, upstream_commit = clone_upstream_template(temp_root)
        upstream_template = upstream_dir / UPSTREAM_TEMPLATE_RELATIVE
        upstream_copying = upstream_dir / UPSTREAM_COPYING_RELATIVE

        package_dir = temp_root / PACKAGE_ROOT
        package_dir.mkdir(parents=True, exist_ok=True)
        package_template = package_dir / "Zotero.dotm"
        shutil.copy2(upstream_template, package_template)

        patch_custom_ui(package_template)
        set_access_vbom(1)
        import_macro_module(package_template, bas_path)
        verify_custom_ui(package_template)
        verify_macro_module(package_template)

        shutil.copy2(mac_install_doc, package_dir / "MAC_INSTALL.md")
        shutil.copy2(mac_install_script, package_dir / "install_mac.command")
        shutil.copy2(mac_restore_script, package_dir / "restore_mac.command")
        shutil.copy2(style_guide, package_dir / "STYLE_GUIDE.md")
        shutil.copy2(upstream_copying, package_dir / "UPSTREAM_COPYING.txt")
        write_upstream_info(package_dir, upstream_commit)
        write_restore_note(package_dir)

        if OUTPUT_ZIP.exists():
            OUTPUT_ZIP.unlink()

        with zipfile.ZipFile(OUTPUT_ZIP, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for path in package_dir.rglob("*"):
                add_file_to_zip(archive, path, str(path.relative_to(temp_root)).replace("\\", "/"))

        return OUTPUT_ZIP
    finally:
        restore_access_vbom(access_vbom_existed, access_vbom_value)
        shutil.rmtree(temp_root, ignore_errors=True)


def main() -> int:
    output = build_package()
    print(f"Built: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
