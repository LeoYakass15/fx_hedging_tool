"""Compile the VBA source files in vba/ into a vbaProject.bin (no Excel needed).

Uses the open-source ms_ovba package (MS-OVBA spec). The project stores the source
code only; Excel compiles it the first time the workbook is opened.

    python tools/build_vba_project.py <out_dir> <sheet codename> [<sheet codename> ...]
"""
import os
import shutil
import sys
import tempfile
import uuid

from ms_ovba.Models.Entities.doc_module import DocModule
from ms_ovba.Models.Entities.reference import Reference
from ms_ovba.Models.Entities.reference_registered import ReferenceRegistered
from ms_ovba.Models.Entities.std_module import StdModule
from ms_ovba.Models.Fields.libid_reference import LibidReference
from ms_ovba.vbaProject import VbaProject
from ms_ovba.Views.project_ole_file import ProjectOleFile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ovba_compress import patch_ms_ovba  # noqa: E402

patch_ms_ovba()

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOC_TEMPLATE = ('VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\n'
                'Attribute VB_Name = "{name}"\r\nAttribute VB_GlobalNameSpace = False\r\n'
                'Attribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\n'
                'Attribute VB_Exposed = True\r\n')
SHEET_GUID = uuid.UUID("0002082000000000C000000000000046")
BOOK_GUID = uuid.UUID("0002081900000000C000000000000046")


def build(out_dir: str, sheets: list[str], workbook_code: str = "") -> str:
    work = tempfile.mkdtemp()
    project = VbaProject()
    project.project_id = '{9E394C0B-697E-4AEE-9FA6-446F51FB30DC}'

    def doc(name, guid, code=""):
        path = os.path.join(work, name + ".cls")
        with open(path, "w", newline="") as f:
            f.write(DOC_TEMPLATE.format(name=name) + code.replace("\r\n", "\n").replace("\n", "\r\n"))
        m = DocModule(name)
        m.add_file(path)
        m.add_guid(guid)
        m.normalize_file()
        project.add_module(m)

    doc("ThisWorkbook", BOOK_GUID, workbook_code)
    for s in sheets:
        doc(s, SHEET_GUID)
    for fname in sorted(os.listdir(os.path.join(ROOT, "vba"))):
        if fname.endswith(".bas"):
            path = os.path.join(work, fname)
            # VBA source must be CRLF / cp1252
            with open(os.path.join(ROOT, "vba", fname), encoding="utf-8") as f:
                txt = f.read().replace("\r\n", "\n").replace("\n", "\r\n")
            with open(path, "w", encoding="cp1252", newline="") as f:
                f.write(txt)
            m = StdModule(os.path.splitext(fname)[0])
            m.add_file(path)
            m.normalize_file()
            project.add_module(m)

    stdole = LibidReference(uuid.UUID("0002043000000000C000000000000046"), "2.0", "0",
                            "C:\\Windows\\System32\\stdole2.tlb", "OLE Automation")
    office = LibidReference(uuid.UUID("2DF8D04C5BFA101BBDE500AA0044DE52"), "2.0", "0",
                            r"C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL",
                            "Microsoft Office 16.0 Object Library")
    project.add_reference(Reference(ReferenceRegistered(stdole), "stdole"))
    project.add_reference(Reference(ReferenceRegistered(office), "Office"))

    cwd = os.getcwd()
    os.chdir(work)
    try:
        ProjectOleFile.write_file(project)
    finally:
        os.chdir(cwd)
    os.makedirs(out_dir, exist_ok=True)
    dst = os.path.join(out_dir, "vbaProject.bin")
    shutil.copy(os.path.join(work, "vbaProject.bin"), dst)
    shutil.rmtree(work, ignore_errors=True)
    return dst


if __name__ == "__main__":
    print(build(sys.argv[1], sys.argv[2:]))
