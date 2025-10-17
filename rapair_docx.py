#!/usr/bin/env python3
"""
repair_docx.py
A robust DOCX repair utility that:
 - backs up the original file
 - inspects core props and styles
 - attempts XML sanitization for document.xml and styles.xml
 - replaces malformed XML if possible
 - converts docx -> markdown -> docx via pandoc as a fallback
 - preserves media and relationships
 - outputs a repair report

Notes:
 - This tool uses python-docx, lxml, pypandoc (pandoc must be installed),
   and standard library modules. See requirements below.
 - Not every corrupted docx can be fully repaired automatically. This script
   attempts multiple safe strategies and logs each step.
"""

import argparse
import zipfile
import shutil
import os
import sys
import tempfile
import datetime
import json
import subprocess
from pathlib import Path
from typing import Tuple, Dict, List

# Third-party
try:
    from docx import Document
except Exception:
    Document = None

try:
    from lxml import etree
except Exception:
    etree = None

try:
    import pypandoc
except Exception:
    pypandoc = None

# ---------------------------
# Requirements note (for README)
# pip install python-docx lxml pypandoc
# pandoc must be installed on system: https://pandoc.org/installing.html
# ---------------------------

# Utility / logging
REPORT_KEYS = [
    "input_path",
    "backup_path",
    "timestamp",
    "actions",
    "errors",
    "final_docx",
    "final_docx_ok",
]

def now_ts():
    return datetime.datetime.utcnow().isoformat() + "Z"

class RepairReport:
    def __init__(self, input_path: str):
        self.data = {k: None for k in REPORT_KEYS}
        self.data["input_path"] = str(input_path)
        self.data["timestamp"] = now_ts()
        self.data["actions"] = []
        self.data["errors"] = []

    def add_action(self, msg: str):
        self.data["actions"].append({"time": now_ts(), "msg": msg})
        print("[ACTION]", msg)

    def add_error(self, msg: str):
        self.data["errors"].append({"time": now_ts(), "msg": msg})
        print("[ERROR]", msg, file=sys.stderr)

    def set(self, key, value):
        self.data[key] = value

    def save(self, path: str):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, indent=2)

# ---------------------------
# Low-level helpers
# ---------------------------

def ensure_pandoc_available(report: RepairReport) -> bool:
    if pypandoc is None:
        report.add_action("pypandoc not installed; pandoc fallback disabled.")
        return False
    try:
        pypandoc.get_pandoc_version()
        report.add_action("pandoc found via pypandoc.")
        return True
    except Exception as e:
        report.add_error(f"pandoc not available: {e}")
        return False

def backup_file(src: Path, report: RepairReport) -> Path:
    dst = src.with_suffix(src.suffix + ".backup." + datetime.datetime.utcnow().strftime("%Y%m%d%H%M%S"))
    shutil.copy2(src, dst)
    report.set("backup_path", str(dst))
    report.add_action(f"Backed up original file to {dst}")
    return dst

def unzip_to_temp(docx_path: Path, report: RepairReport) -> Path:
    tmp = Path(tempfile.mkdtemp(prefix="docx_repair_"))
    report.add_action(f"Created temp dir {tmp}")
    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            z.extractall(tmp)
        report.add_action("Extracted docx zip to temp dir")
        return tmp
    except Exception as e:
        report.add_error(f"Failed to unzip {docx_path}: {e}")
        raise

def rezip_from_temp(tmpdir: Path, out_docx: Path, report: RepairReport):
    # Create zip in the correct order (optional) - simple approach works usually
    with zipfile.ZipFile(out_docx, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(tmpdir):
            for fname in files:
                fpath = Path(root) / fname
                arcname = str(fpath.relative_to(tmpdir)).replace("\\", "/")
                z.write(fpath, arcname)
    report.add_action(f"Rebuilt docx as {out_docx}")

def safe_parse_xml(path: Path, report: RepairReport) -> Tuple[bool, str]:
    """
    Try to parse the file with lxml and return (ok, pretty_xml_or_error)
    """
    if etree is None:
        report.add_error("lxml not installed; cannot parse XML safely.")
        return False, "lxml missing"
    try:
        parser = etree.XMLParser(recover=True, remove_blank_text=True)
        tree = etree.parse(str(path), parser=parser)
        pretty = etree.tostring(tree, encoding='utf-8', xml_declaration=True, pretty_print=True).decode('utf-8')
        return True, pretty
    except Exception as e:
        report.add_error(f"XML parse error for {path}: {e}")
        return False, str(e)

def write_text_file(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)

# ---------------------------
# High-level repair strategies
# ---------------------------

def fix_core_properties(tmpdir: Path, report: RepairReport):
    """
    Inspect /docProps/core.xml and modify common missing tags if present.
    """
    core_path = tmpdir / "docProps" / "core.xml"
    if not core_path.exists():
        report.add_action("No core.xml found - nothing to fix for metadata.")
        return
    ok, content = safe_parse_xml(core_path, report)
    if not ok:
        report.add_action("Attempting to rewrite core.xml with safer template.")
        # Attempt to create a minimal core.xml
        minimal = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Repaired Document</dc:title>
  <dc:creator>AutoRepair</dc:creator>
  <cp:lastModifiedBy>AutoRepair</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{datetime.datetime.utcnow().isoformat()}Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{datetime.datetime.utcnow().isoformat()}Z</dcterms:modified>
</cp:coreProperties>'''
        write_text_file(core_path, minimal)
        report.add_action("Wrote minimal core.xml")
    else:
        # If parsed ok, ensure basic tags are present
        try:
            root = etree.fromstring(content.encode('utf-8'))
            ns = {'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                  'dc': 'http://purl.org/dc/elements/1.1/'}
            title = root.find('.//{http://purl.org/dc/elements/1.1/}title')
            creator = root.find('.//{http://purl.org/dc/elements/1.1/}creator')
            if title is None or (title.text or "").strip() == "":
                # insert title element
                title_el = etree.Element("{http://purl.org/dc/elements/1.1/}title")
                title_el.text = "Repaired Document"
                root.insert(0, title_el)
                report.add_action("Inserted missing title tag in core.xml")
            if creator is None:
                creator_el = etree.Element("{http://purl.org/dc/elements/1.1/}creator")
                creator_el.text = "AutoRepair"
                root.insert(0, creator_el)
                report.add_action("Inserted missing creator tag in core.xml")
            pretty = etree.tostring(root, encoding='utf-8', xml_declaration=True, pretty_print=True).decode('utf-8')
            write_text_file(core_path, pretty)
            report.add_action("Rewrote core.xml with ensured basic metadata")
        except Exception as e:
            report.add_error(f"Failed to rewrite core.xml properly: {e}")

def sanitize_xml_files(tmpdir: Path, report: RepairReport, target_files: List[str]):
    """
    For each `target_files` under tmpdir (relative paths), try parsing & pretty-printing.
    If parsing fails, attempt to recover via lxml's recover mode and overwrite.
    """
    for rel in target_files:
        path = tmpdir / rel
        if not path.exists():
            report.add_action(f"{rel} not present; skipping.")
            continue
        ok, content = safe_parse_xml(path, report)
        if ok:
            # Overwrite with pretty content
            try:
                write_text_file(path, content)
                report.add_action(f"Sanitized XML for {rel}")
            except Exception as e:
                report.add_error(f"Failed to write sanitized XML to {rel}: {e}")
        else:
            report.add_action(f"Could not fully parse {rel}; leaving for pandoc fallback.")

def remove_custom_xml(tmpdir: Path, report: RepairReport):
    """
    Remove customXml/ if it contains invalid items that sometimes break Word.
    We keep a backup copy inside temp for debugging if removed.
    """
    custom_dir = tmpdir / "customXml"
    if custom_dir.exists():
        backup_dir = tmpdir / "customXml.removed"
        shutil.move(str(custom_dir), str(backup_dir))
        report.add_action("Moved customXml to customXml.removed (some Word versions choke on custom XML).")

def try_open_with_python_docx(docx_path: Path, report: RepairReport) -> bool:
    """
    Try to open with python-docx to validate the docx. Returns True if ok.
    """
    if Document is None:
        report.add_error("python-docx not installed; cannot validate with Document().")
        return False
    try:
        _ = Document(str(docx_path))
        report.add_action("python-docx successfully opened the file (basic validation OK).")
        return True
    except Exception as e:
        report.add_error(f"python-docx failed to open file: {e}")
        return False

def pandoc_roundtrip(tmpdir: Path, repaired_docx_path: Path, report: RepairReport) -> bool:
    """
    Do docx -> markdown -> docx using pandoc as a robust fallback to rebuild formatting.
    Requires pandoc to be installed.
    """
    if not ensure_pandoc_available(report):
        return False
    # find the original file inside tmpdir, but pandoc wants an actual .docx file.
    # We'll create a temp docx from tmpdir rezip
    tmp_rezip = repaired_docx_path.with_name(repaired_docx_path.stem + ".pandoc_temp.docx")
    rezip_from_temp(tmpdir, tmp_rezip, report)
    md = tmp_rezip.with_suffix(".md")
    try:
        report.add_action("Running pandoc docx -> md")
        pypandoc.convert_file(str(tmp_rezip), 'md', outputfile=str(md))
        report.add_action(f"Pandoc conversion to markdown saved at {md}")
        # Now create a clean docx from md
        new_docx = repaired_docx_path.with_name(repaired_docx_path.stem + ".pandoc_rebuilt.docx")
        report.add_action("Running pandoc md -> docx to rebuild structure")
        pypandoc.convert_file(str(md), 'docx', outputfile=str(new_docx))
        if new_docx.exists():
            shutil.copy2(new_docx, repaired_docx_path)
            report.add_action(f"Pandoc rebuilt document replaced repaired docx: {repaired_docx_path}")
            return True
    except Exception as e:
        report.add_error(f"Pandoc roundtrip failed: {e}")
    finally:
        # cleanup temp
        for f in [tmp_rezip, md]:
            try:
                if f.exists(): f.unlink()
            except Exception:
                pass
    return False

# ---------------------------
# Orchestration
# ---------------------------

def repair_docx(input_path: Path, output_path: Path = None, verbose: bool = True) -> RepairReport:
    report = RepairReport(input_path=str(input_path))
    try:
        # Sanity check input exists
        if not input_path.exists():
            report.add_error(f"Input file not found: {input_path}")
            return report

        # Backup original
        bak = backup_file(input_path, report)

        # Unzip to temp dir
        tmpdir = unzip_to_temp(input_path, report)

        # Basic fixes
        fix_core_properties(tmpdir, report)
        sanitize_xml_files(tmpdir, report, ["word/document.xml", "word/styles.xml", "word/_rels/document.xml.rels"])
        remove_custom_xml(tmpdir, report)

        # Rebuild docx
        repaired_docx = input_path.with_name(input_path.stem + ".repaired.docx") if output_path is None else output_path
        rezip_from_temp(tmpdir, repaired_docx, report)

        report.add_action("Attempting to validate repaired docx with python-docx.")
        ok = try_open_with_python_docx(repaired_docx, report)

        if not ok:
            report.add_action("python-docx validation failed - attempting pandoc roundtrip fallback.")
            pandoc_ok = pandoc_roundtrip(tmpdir, repaired_docx, report)
            if pandoc_ok:
                report.add_action("Pandoc fallback succeeded.")
            else:
                report.add_error("Pandoc fallback also failed. Final docx may still be corrupt.")
        else:
            report.add_action("Repaired document validated successfully with python-docx.")

        report.set("final_docx", str(repaired_docx))
        report.set("final_docx_ok", ok)

        # Save final report into same folder
        report_file = input_path.with_suffix(".repair_report.json")
        report.save(str(report_file))
        report.add_action(f"Saved repair report to {report_file}")

        # clean temp
        try:
            shutil.rmtree(tmpdir)
            report.add_action(f"Removed temp dir {tmpdir}")
        except Exception as e:
            report.add_error(f"Failed to remove temp dir {tmpdir}: {e}")

    except Exception as e:
        report.add_error(f"Unexpected error during repair: {e}")

    return report

# ---------------------------
# CLI entry
# ---------------------------

def parse_args():
    parser = argparse.ArgumentParser(description="Repair a corrupt DOCX by sanitizing XML and rebuilding via pandoc.")
    parser.add_argument("input", type=str, help="Path to input DOCX file (corrupted.docx)")
    parser.add_argument("-o", "--output", type=str, default=None, help="Optional path to save repaired DOCX")
    parser.add_argument("-q", "--quiet", action="store_true", help="Quiet mode")
    return parser.parse_args()

def main():
    args = parse_args()
    inp = Path(args.input)
    out = Path(args.output) if args.output else None
    report = repair_docx(inp, out, verbose=not args.quiet)
    summary = {
        "input": report.data.get("input_path"),
        "backup": report.data.get("backup_path"),
        "final": report.data.get("final_docx"),
        "ok": report.data.get("final_docx_ok"),
        "errors": report.data.get("errors"),
        "actions": report.data.get("actions")[-5:],
    }
    print("\n=== Repair Summary ===")
    print(json.dumps(summary, indent=2))

if __name__ == "__main__":
    main()

