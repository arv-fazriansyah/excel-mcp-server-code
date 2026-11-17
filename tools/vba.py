"""VBA operations (supports .xlsm/.xlsb/.xls)."""
from __future__ import annotations

import hashlib
import os
import re
from pathlib import Path

from session import get_session
from com_executor import get_executor
from config import ALLOW_VBA_WRITE, ALLOW_MACROS, MACRO_EXTS


VB_TYPE = {1: "Module", 2: "Class", 3: "Form", 100: "Document"}
VB_EXT = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}


def _clean_module_source(text: str) -> str:
    out = []
    in_begin = False
    for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        stripped = line.strip()
        if re.match(r"^VERSION\s+[\d.]+\s+CLASS$", stripped, re.IGNORECASE):
            continue
        if re.match(r"^BEGIN\b", stripped, re.IGNORECASE):
            in_begin = True
            continue
        if in_begin:
            if re.match(r"^END$", stripped, re.IGNORECASE):
                in_begin = False
            continue
        if re.match(r"^\s*Attribute\s+VB_\w+\s*=", line, re.IGNORECASE):
            continue
        if re.match(r"^MultiUse\s*=", stripped, re.IGNORECASE):
            continue
        out.append(line)
    while out and not out[0].strip():
        out.pop(0)
    return "\n".join(out).rstrip() + "\n"


def _ensure_vba_file(path: str):
    ext = os.path.splitext(path)[1].lower()
    if ext not in MACRO_EXTS:
        raise ValueError(f"VBA operations require macro-enabled workbooks (.xlsm/.xlsb/.xls). Current: {ext}")


def _vb_project(wb):
    try:
        return wb.VBProject
    except Exception as exc:
        raise PermissionError("Akses VBA project ditolak. Aktifkan 'Trust access to the VBA project object model' di Excel Trust Center. Detail: " + str(exc))


def vba_status(sid: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    ex = get_executor()
    _ensure_vba_file(s.path)
    def _run():
        try:
            vp = wb.VBProject
            access = True
            access_msg = None
        except Exception as exc:
            vp = None
            access = False
            access_msg = str(exc)
        protection = None
        mode = None
        if vp is not None:
            try:
                protection = "Locked" if vp.Protection == 1 else "None"
            except Exception:
                protection = None
            try:
                mode = {0: "Design", 1: "Run", 2: "Break"}.get(int(vp.Mode), str(vp.Mode))
            except Exception:
                mode = None
        return {"path": s.path, "projectAccess": access, "accessMessage": access_msg,
                "protection": protection, "mode": mode, "macrosAllowed": ALLOW_MACROS}
    return ex.call(_run)


def vba_list_modules(sid: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    _ensure_vba_file(s.path)
    ex = get_executor()
    def _run():
        vp = _vb_project(wb)
        modules = []
        for comp in vp.VBComponents:
            cm = comp.CodeModule
            code = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
            procs = []
            for m in re.finditer(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)", code, re.IGNORECASE | re.MULTILINE):
                kind = m.group(1).split()[0].lower()
                procs.append({"name": m.group(2), "kind": kind,
                              "startLine": code[:m.start()].count("\n") + 1})
            modules.append({"name": comp.Name, "type": VB_TYPE.get(comp.Type, str(comp.Type)),
                            "lineCount": cm.CountOfLines, "procedures": [p["name"] for p in procs],
                            "procedureDetails": procs})
        return {"modules": modules}
    return ex.call(_run)


def vba_read_module(sid: str, module_name: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    _ensure_vba_file(s.path)
    ex = get_executor()
    def _run():
        comp = _vb_project(wb).VBComponents(module_name)
        cm = comp.CodeModule
        code = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
        return {"module": module_name, "type": VB_TYPE.get(comp.Type, str(comp.Type)),
                "lineCount": cm.CountOfLines, "source": code}
    return ex.call(_run)


def vba_read_procedure(sid: str, module_name: str, procedure_name: str, procedure_kind: str = "Sub",
                       start_line: int | None = None, line_count: int | None = None) -> dict:
    s = get_session(sid)
    wb = s.workbook
    _ensure_vba_file(s.path)
    ex = get_executor()
    def _run():
        comp = _vb_project(wb).VBComponents(module_name)
        cm = comp.CodeModule
        full = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
        # find procedure
        if start_line is not None:
            sl = start_line
            cnt = line_count or 500
            chunk = cm.Lines(sl, min(cnt, cm.CountOfLines - sl + 1)) if cm.CountOfLines >= sl else ""
            return {"module": module_name, "startLine": sl, "returnedLineCount": len(chunk.splitlines()),
                    "hasMore": sl + len(chunk.splitlines()) - 1 < cm.CountOfLines,
                    "nextStartLine": sl + len(chunk.splitlines()),
                    "source": chunk}
        # procedure by name
        pat = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
                         + re.escape(procedure_kind) + r"\s+" + re.escape(procedure_name) + r"\b",
                         re.IGNORECASE | re.MULTILINE)
        m = pat.search(full)
        if not m:
            raise ValueError(f"Prosedur {procedure_name} ({procedure_kind}) tidak ditemukan di {module_name}")
        start = full[:m.start()].count("\n") + 1
        # find matching End
        lines = full.split("\n")
        # procedure occupies from start line to the End <kind> line
        end_kind = procedure_kind.split()[0].lower()
        end_re = re.compile(r"^\s*End\s+" + end_kind + r"\b", re.IGNORECASE)
        end_idx = None
        for i in range(start - 1, len(lines)):
            if end_re.match(lines[i]):
                end_idx = i
                break
        if end_idx is None:
            end_idx = len(lines) - 1
        chunk_lines = lines[start - 1:end_idx + 1]
        chunk = "\n".join(chunk_lines)
        # introductory comments: walk back from start while previous line is comment/blank
        intro = start
        i = start - 2
        while i >= 0 and (lines[i].strip() == "" or lines[i].strip().startswith("'")):
            intro = i + 1
            i -= 1
        body_start = start
        # adjust: start already at declaration; intro includes comments above
        source_hash = hashlib.sha256(chunk.encode("utf-8")).hexdigest()
        return {"module": module_name, "procedure": procedure_name, "kind": procedure_kind,
                "startLine": intro, "bodyStartLine": body_start,
                "returnedLineCount": len(chunk_lines), "source": chunk,
                "sourceHash": source_hash, "hasMore": False}
    return ex.call(_run)


def vba_search(sid: str, search_text: str, module_name: str | None = None, whole_word: bool = False,
               match_case: bool = False, max_matches: int = 50) -> dict:
    s = get_session(sid)
    wb = s.workbook
    _ensure_vba_file(s.path)
    ex = get_executor()
    def _run():
        vp = _vb_project(wb)
        comps = [vp.VBComponents(module_name)] if module_name else list(vp.VBComponents)
        matches = []
        flags = 0 if match_case else re.IGNORECASE
        pattern = re.escape(search_text)
        if whole_word:
            pattern = r"\b" + pattern + r"\b"
        rx = re.compile(pattern, flags)
        for comp in comps:
            cm = comp.CodeModule
            code = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
            for m in rx.finditer(code):
                line_no = code[:m.start()].count("\n") + 1
                line_text = code.split("\n")[line_no - 1] if line_no - 1 < len(code.split("\n")) else ""
                excerpt = line_text.strip()[:200]
                matches.append({"module": comp.Name, "line": line_no, "column": m.start() - code.rfind("\n", 0, m.start()) ,
                                "excerpt": excerpt})
                if len(matches) >= max_matches:
                    break
            if len(matches) >= max_matches:
                break
        return {"matches": matches, "count": len(matches), "hasMore": len(matches) >= max_matches}
    return ex.call(_run)


def vba_write_module(sid: str, module_name: str, code: str, create: bool = False) -> dict:
    if not ALLOW_VBA_WRITE:
        raise PermissionError("Tulis VBA dinonaktifkan. Set EXCEL_MCP_ALLOW_VBA_WRITE=1 lalu restart.")
    s = get_session(sid)
    wb = s.workbook
    _ensure_vba_file(s.path)
    ex = get_executor()
    def _run():
        vp = _vb_project(wb)
        try:
            comp = vp.VBComponents(module_name)
            created = False
        except Exception:
            if not create:
                raise ValueError(f"Module '{module_name}' tidak ada. Pakai create=true.")
            comp = vp.VBComponents.Add(1)
            comp.Name = module_name
            created = True
        clean = _clean_module_source(code)
        cm = comp.CodeModule
        if cm.CountOfLines:
            cm.DeleteLines(1, cm.CountOfLines)
        if clean.strip():
            cm.AddFromString(clean)
        stored = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
        return {"module": module_name, "created": created, "lines": cm.CountOfLines,
                "exactMatch": stored.replace("\r\n", "\n").rstrip() == clean.replace("\r\n", "\n").rstrip()}
    return ex.call(_run)


def vba_import_module(sid: str, module_name: str, code: str | None = None, file_path: str | None = None) -> dict:
    if not ALLOW_VBA_WRITE:
        raise PermissionError("Tulis VBA dinonaktifkan. Set EXCEL_MCP_ALLOW_VBA_WRITE=1 lalu restart.")
    if file_path:
        from utils import check_path
        check_path(file_path)
        code = Path(file_path).read_text(encoding="utf-8-sig")
    if code is None:
        raise ValueError("Butuh code atau file_path")
    return vba_write_module(sid, module_name, code, create=True)


def vba_delete_module(sid: str, module_name: str) -> dict:
    if not ALLOW_VBA_WRITE:
        raise PermissionError("Hapus VBA dinonaktifkan. Set EXCEL_MCP_ALLOW_VBA_WRITE=1 lalu restart.")
    s = get_session(sid)
    wb = s.workbook
    _ensure_vba_file(s.path)
    ex = get_executor()
    def _run():
        vp = _vb_project(wb)
        comp = vp.VBComponents(module_name)
        vp.VBComponents.Remove(comp)
        return {"deleted": module_name}
    return ex.call(_run)


def vba_export_module(sid: str, module_name: str | None, out_dir: str) -> dict:
    from utils import check_path
    check_path(out_dir)
    s = get_session(sid)
    wb = s.workbook
    _ensure_vba_file(s.path)
    os.makedirs(out_dir, exist_ok=True)
    ex = get_executor()
    def _run():
        vp = _vb_project(wb)
        comps = [vp.VBComponents(module_name)] if module_name else list(vp.VBComponents)
        files = []
        for comp in comps:
            ext = VB_EXT.get(comp.Type, ".txt")
            target = os.path.join(out_dir, comp.Name + ext)
            comp.Export(target)
            files.append(target)
        return {"outDir": out_dir, "files": files}
    return ex.call(_run)


def vba_replace_procedure(sid: str, module_name: str, procedure_name: str, procedure_kind: str,
                          expected_source_hash: str, code: str) -> dict:
    if not ALLOW_VBA_WRITE:
        raise PermissionError("Tulis VBA dinonaktifkan. Set EXCEL_MCP_ALLOW_VBA_WRITE=1 lalu restart.")
    # read current
    current = vba_read_procedure(sid, module_name, procedure_name, procedure_kind)
    if current.get("sourceHash") != expected_source_hash:
        raise ValueError("Source berubah sejak dibaca. Baca ulang sebelum mengedit.")
    s = get_session(sid)
    wb = s.workbook
    _ensure_vba_file(s.path)
    ex = get_executor()
    def _run():
        comp = _vb_project(wb).VBComponents(module_name)
        cm = comp.CodeModule
        full = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
        lines = full.split("\n")
        # find procedure span
        pat = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?" + re.escape(procedure_kind) + r"\s+" + re.escape(procedure_name) + r"\b", re.IGNORECASE)
        start_idx = None
        for i, ln in enumerate(lines):
            if pat.match(ln):
                start_idx = i
                break
        if start_idx is None:
            raise ValueError("Prosedur tidak ditemukan")
        end_kind = procedure_kind.split()[0].lower()
        end_re = re.compile(r"^\s*End\s+" + end_kind + r"\b", re.IGNORECASE)
        end_idx = None
        for i in range(start_idx, len(lines)):
            if end_re.match(lines[i]):
                end_idx = i
                break
        if end_idx is None:
            raise ValueError("End prosedur tidak ditemukan")
        # preserve intro comments above
        intro_idx = start_idx
        i = start_idx - 1
        while i >= 0 and (lines[i].strip() == "" or lines[i].strip().startswith("'")):
            intro_idx = i
            i -= 1
        new_lines = lines[:intro_idx] + code.strip().split("\n") + lines[end_idx + 1:]
        new_code = "\n".join(new_lines)
        if cm.CountOfLines:
            cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(new_code)
        new_proc = vba_read_procedure(sid, module_name, procedure_name, procedure_kind)
        return {"module": module_name, "procedure": procedure_name, "newSourceHash": new_proc["sourceHash"]}
    return ex.call(_run)


def vba_run_macro(sid: str, macro: str, args: list | None = None, timeout_seconds: int = 120) -> dict:
    if not ALLOW_MACROS:
        raise PermissionError("Eksekusi macro dinonaktifkan. Set EXCEL_MCP_ALLOW_MACROS=1 lalu restart.")
    s = get_session(sid)
    wb = s.workbook
    _ensure_vba_file(s.path)
    ex = get_executor()
    call_args = args or []
    target = macro if "!" in macro else f"'{wb.Name}'!{macro}"
    def _run():
        return ex.call(lambda: s.app.Run(target, *call_args))
    try:
        res = _run()
    except Exception as exc:
        raise RuntimeError(f"Macro '{macro}' gagal: {exc}")
    return {"macro": target, "result": res}


def vba_references(sid: str) -> dict:
    s = get_session(sid)
    wb = s.workbook
    _ensure_vba_file(s.path)
    ex = get_executor()
    def _run():
        vp = _vb_project(wb)
        refs = []
        try:
            for r in vp.References:
                entry = {"index": r.Index, "isBroken": bool(r.IsBroken)}
                if not r.IsBroken:
                    entry.update({"name": r.Name, "description": r.Description,
                                  "libraryId": r.Guid, "major": r.Major, "minor": r.Minor,
                                  "builtIn": bool(r.BuiltIn)})
                refs.append(entry)
        except Exception as exc:
            raise RuntimeError(f"Gagal membaca references: {exc}")
        return {"references": refs, "hasBrokenReferences": any(r.get("isBroken") for r in refs)}
    return ex.call(_run)


def vba_userform_create(sid: str, name: str) -> dict:
    if not ALLOW_VBA_WRITE:
        raise PermissionError("Tulis VBA dinonaktifkan.")
    s = get_session(sid)
    wb = s.workbook
    _ensure_vba_file(s.path)
    ex = get_executor()
    def _run():
        vp = _vb_project(wb)
        comp = vp.VBComponents.Add(3)  # vbext_ct_MSForm
        comp.Name = name
        return {"name": name, "type": "Form"}
    return ex.call(_run)


def vba_userform_add_control(sid: str, form_name: str, control_type: str, name: str,
                             left: float = 10, top: float = 10, width: float = 80, height: float = 20) -> dict:
    if not ALLOW_VBA_WRITE:
        raise PermissionError("Tulis VBA dinonaktifkan.")
    # control_type: CommandButton, TextBox, Label, ComboBox, ListBox, CheckBox, OptionButton, Frame, Image
    PROGIDS = {"CommandButton": "Forms.CommandButton.1", "TextBox": "Forms.TextBox.1",
               "Label": "Forms.Label.1", "ComboBox": "Forms.ComboBox.1",
               "ListBox": "Forms.ListBox.1", "CheckBox": "Forms.CheckBox.1",
               "OptionButton": "Forms.OptionButton.1", "Frame": "Forms.Frame.1",
               "Image": "Forms.Image.1"}
    if control_type not in PROGIDS:
        raise ValueError(f"Tipe kontrol tidak dikenal: {control_type}")
    s = get_session(sid)
    wb = s.workbook
    _ensure_vba_file(s.path)
    ex = get_executor()
    def _run():
        vp = _vb_project(wb)
        form = vp.VBComponents(form_name)
        ctrl = form.Designer.Controls.Add(PROGIDS[control_type], name, True)
        ctrl.Left, ctrl.Top, ctrl.Width, ctrl.Height = left, top, width, height
        return {"form": form_name, "control": name, "type": control_type}
    return ex.call(_run)
