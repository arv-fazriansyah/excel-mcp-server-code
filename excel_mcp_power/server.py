\"\"\"FastMCP server for Excel MCP Power.\"\"\"
from __future__ import annotations

import os
from mcp.server.fastmcp import FastMCP

from .config import ROOT, READONLY, ALLOW_MACROS, ALLOW_VBA_WRITE
from .session import (
    session_open, session_create, session_attach, session_info,
    session_save, session_save_as, session_close, session_list
)
from .tools import sheet, range, vba, workbook, table, named

mcp = FastMCP("excel-mcp-power")

# --- Session Tools ---
@mcp.tool()
def open_workbook(path: str, read_only: bool = False, password: str | None = None) -> dict:
    \"\"\"Membuka workbook Excel (mendukung .xlsx, .xlsm, .xlsb, .xls).\"\"\"
    return session_open(path, read_only=read_only or READONLY, password=password)

@mcp.tool()
def create_workbook(path: str, sheets: str = "", force: bool = False) -> dict:
    \"\"\"Membuat workbook baru dengan daftar sheet (pisahkan koma).\"\"\"
    return session_create(path, sheets=sheets, force=force)

@mcp.tool()
def attach_workbook() -> dict:
    \"\"\"Melihat daftar workbook yang sedang terbuka di instance Excel.\"\"\"
    return session_attach()

@mcp.tool()
def workbook_info(sessionId: str) -> dict:
    \"\"\"Mendapatkan informasi detail session dan struktur workbook.\"\"\"
    return session_info(sessionId)

@mcp.tool()
def save_workbook(sessionId: str) -> dict:
    \"\"\"Menyimpan workbook aktif di session.\"\"\"
    return session_save(sessionId)

@mcp.tool()
def save_workbook_as(sessionId: str, targetPath: str, format: str = "auto", overwrite: bool = False) -> dict:
    \"\"\"Menyimpan workbook ke path baru dengan format tertentu (.xlsx, .xlsm, .xlsb).\"\"\"
    return session_save_as(sessionId, targetPath, fmt=format, overwrite=overwrite)

@mcp.tool()
def close_workbook(sessionId: str, save: bool = False) -> dict:
    \"\"\"Menutup session workbook.\"\"\"
    return session_close(sessionId, save=save)

@mcp.tool()
def list_sessions() -> dict:
    \"\"\"Mendaftarkan seluruh session aktif.\"\"\"
    return session_list()


# --- Sheet Tools ---
@mcp.tool()
def list_sheets(sessionId: str) -> dict:
    \"\"\"Mendaftarkan semua worksheet dalam workbook.\"\"\"
    return sheet.sheet_list(sessionId)

@mcp.tool()
def add_sheet(sessionId: str, name: str | None = None, after: str | None = None) -> dict:
    \"\"\"Menambahkan worksheet baru.\"\"\"
    return sheet.sheet_add(sessionId, name=name, after=after)

@mcp.tool()
def rename_sheet(sessionId: str, oldName: str, newName: str) -> dict:
    \"\"\"Mengubah nama worksheet.\"\"\"
    return sheet.sheet_rename(sessionId, oldName, newName)

@mcp.tool()
def delete_sheet(sessionId: str, name: str) -> dict:
    \"\"\"Menghapus worksheet.\"\"\"
    return sheet.sheet_delete(sessionId, name)


# --- Range Tools ---
@mcp.tool()
def get_range_values(sessionId: str, sheet: str, rangeAddress: str) -> dict:
    \"\"\"Membaca nilai dari range sel (mengembalikan 2D grid).\"\"\"
    return range.range_get_values(sessionId, sheet, rangeAddress)

@mcp.tool()
def set_range_values(sessionId: str, sheet: str, rangeAddress: str, values: list) -> dict:
    \"\"\"Menulis nilai (2D grid atau skalar) ke range sel.\"\"\"
    return range.range_set_values(sessionId, sheet, rangeAddress, values)

@mcp.tool()
def get_range_formulas(sessionId: str, sheet: str, rangeAddress: str) -> dict:
    \"\"\"Membaca formula dari range sel.\"\"\"
    return range.range_get_formulas(sessionId, sheet, rangeAddress)

@mcp.tool()
def set_range_formula(sessionId: str, sheet: str, rangeAddress: str, formula: str) -> dict:
    \"\"\"Menulis formula ke range sel.\"\"\"
    return range.range_set_formula(sessionId, sheet, rangeAddress, formula)

@mcp.tool()
def clear_range(sessionId: str, sheet: str, rangeAddress: str, contentsOnly: bool = False) -> dict:
    \"\"\"Membersihkan isi atau format range sel.\"\"\"
    return range.range_clear(sessionId, sheet, rangeAddress, contents_only=contentsOnly)

@mcp.tool()
def find_in_range(sessionId: str, sheet: str, rangeAddress: str, what: str, matchCase: bool = False, wholeCell: bool = False) -> dict:
    \"\"\"Mencari teks/nilai dalam range.\"\"\"
    return range.range_find(sessionId, sheet, rangeAddress, what, match_case=matchCase, whole_cell=wholeCell)


# --- VBA Tools ---
@mcp.tool()
def vba_status(sessionId: str) -> dict:
    \"\"\"Mengecek status VBA project dan akses VBOM.\"\"\"
    return vba.vba_status(sessionId)

@mcp.tool()
def vba_list_modules(sessionId: str) -> dict:
    \"\"\"Mendaftarkan semua modul VBA dan prosedur di dalamnya.\"\"\"
    return vba.vba_list_modules(sessionId)

@mcp.tool()
def vba_read_module(sessionId: str, moduleName: str) -> dict:
    \"\"\"Membaca source code lengkap suatu modul VBA.\"\"\"
    return vba.vba_read_module(sessionId, moduleName)

@mcp.tool()
def vba_write_module(sessionId: str, moduleName: str, code: str, create: bool = False) -> dict:
    \"\"\"Menulis atau memperbarui source code modul VBA.\"\"\"
    return vba.vba_write_module(sessionId, moduleName, code, create=create)

@mcp.tool()
def vba_run_macro(sessionId: str, macro: str, args: list | None = None) -> dict:
    \"\"\"Menjalankan macro / prosedur VBA.\"\"\"
    return vba.vba_run_macro(sessionId, macro, args=args)


# --- Table & Named Range Tools ---
@mcp.tool()
def list_tables(sessionId: str, sheet: str) -> dict:
    \"\"\"Mendaftarkan Excel Tables (ListObjects) dalam sheet.\"\"\"
    return table.table_list(sessionId, sheet)

@mcp.tool()
def add_table(sessionId: str, sheet: str, rangeAddress: str, name: str | None = None, hasHeader: bool = True) -> dict:
    \"\"\"Membuat Excel Table dari range.\"\"\"
    return table.table_add(sessionId, sheet, rangeAddress, name=name, has_header=hasHeader)

@mcp.tool()
def list_named_ranges(sessionId: str) -> dict:
    \"\"\"Mendaftarkan semua Named Ranges di workbook.\"\"\"
    return named.named_range_list(sessionId)

@mcp.tool()
def add_named_range(sessionId: str, name: str, refersTo: str, sheet: str | None = None) -> dict:
    \"\"\"Membuat atau memperbarui Named Range.\"\"\"
    return named.named_range_add(sessionId, name, refersTo, sheet=sheet)


def main():
    mcp.run()

if __name__ == "__main__":
    main()
