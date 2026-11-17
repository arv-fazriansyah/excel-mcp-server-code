\"\"\"FastMCP server for Excel MCP Power (Consolidated).\"\"\"
from __future__ import annotations

from mcp.server.fastmcp import FastMCP
from excel_engine import (
    session_open, session_create, session_attach, session_info,
    session_save, session_close, session_list,
    sheet_list, sheet_add, sheet_rename, sheet_delete,
    range_get_values, range_set_values, range_get_formulas, range_set_formula,
    table_list, table_add, named_range_list, named_range_add,
    vba_status, vba_list_modules, vba_read_module, vba_run_macro,
    optimize_ribbon_developer
)

mcp = FastMCP("excel-mcp-power")

@mcp.tool()
def open_workbook(path: str, read_only: bool = False, password: str | None = None) -> dict:
    \"\"\"Membuka workbook Excel (.xlsx, .xlsm, .xlsb, .xls).\"\"\"
    return session_open(path, read_only=read_only, password=password)

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
def close_workbook(sessionId: str, save: bool = False) -> dict:
    \"\"\"Menutup session workbook.\"\"\"
    return session_close(sessionId, save=save)

@mcp.tool()
def list_sessions() -> dict:
    \"\"\"Mendaftarkan seluruh session aktif.\"\"\"
    return session_list()

@mcp.tool()
def list_sheets(sessionId: str) -> dict:
    \"\"\"Mendaftarkan semua worksheet dalam workbook.\"\"\"
    return sheet_list(sessionId)

@mcp.tool()
def add_sheet(sessionId: str, name: str | None = None, after: str | None = None) -> dict:
    \"\"\"Menambahkan worksheet baru.\"\"\"
    return sheet_add(sessionId, name=name, after=after)

@mcp.tool()
def rename_sheet(sessionId: str, oldName: str, newName: str) -> dict:
    \"\"\"Mengubah nama worksheet.\"\"\"
    return sheet_rename(sessionId, oldName, newName)

@mcp.tool()
def delete_sheet(sessionId: str, name: str) -> dict:
    \"\"\"Menghapus worksheet.\"\"\"
    return sheet_delete(sessionId, name)

@mcp.tool()
def get_range_values(sessionId: str, sheet: str, rangeAddress: str) -> dict:
    \"\"\"Membaca nilai dari range sel (mengembalikan 2D grid).\"\"\"
    return range_get_values(sessionId, sheet, rangeAddress)

@mcp.tool()
def set_range_values(sessionId: str, sheet: str, rangeAddress: str, values: list) -> dict:
    \"\"\"Menulis nilai (2D grid atau skalar) ke range sel.\"\"\"
    return range_set_values(sessionId, sheet, rangeAddress, values)

@mcp.tool()
def get_range_formulas(sessionId: str, sheet: str, rangeAddress: str) -> dict:
    \"\"\"Membaca formula dari range sel.\"\"\"
    return range_get_formulas(sessionId, sheet, rangeAddress)

@mcp.tool()
def set_range_formula(sessionId: str, sheet: str, rangeAddress: str, formula: str) -> dict:
    \"\"\"Menulis formula ke range sel.\"\"\"
    return range_set_formula(sessionId, sheet, rangeAddress, formula)

@mcp.tool()
def list_tables(sessionId: str, sheet: str) -> dict:
    \"\"\"Mendaftarkan Excel Tables (ListObjects) dalam sheet.\"\"\"
    return table_list(sessionId, sheet)

@mcp.tool()
def add_table(sessionId: str, sheet: str, rangeAddress: str, name: str | None = None, hasHeader: bool = True) -> dict:
    \"\"\"Membuat Excel Table dari range.\"`\"\"
    return table_add(sessionId, sheet, rangeAddress, name=name, has_header=hasHeader)

@mcp.tool()
def list_named_ranges(sessionId: str) -> dict:
    \"\"\"Mendaftarkan semua Named Ranges di workbook.\"\"\"
    return named_range_list(sessionId)

@mcp.tool()
def add_named_range(sessionId: str, name: str, refersTo: str) -> dict:
    \"\"\"Membuat atau memperbarui Named Range.\"\"\"
    return named_range_add(sessionId, name, refersTo)

@mcp.tool()
def vba_project_status(sessionId: str) -> dict:
    \"\"\"Mengecek status VBA project dan akses VBOM.\"\"\"
    return vba_status(sessionId)

@mcp.tool()
def vba_list_modules_tool(sessionId: str) -> dict:
    \"\"\"Mendaftarkan semua modul VBA.\"\"\"
    return vba_list_modules(sessionId)

@mcp.tool()
def vba_read_module_source(sessionId: str, moduleName: str) -> dict:
    \"\"\"Membaca source code lengkap suatu modul VBA.\"\"\"
    return vba_read_module(sessionId, moduleName)

@mcp.tool()
def vba_run_macro_tool(sessionId: str, macro: str, args: list | None = None) -> dict:
    \"\"\"Menjalankan macro / prosedur VBA.\"\"\"
    return vba_run_macro(sessionId, macro, args=args)

@mcp.tool()
def enable_developer_tab(filePath: str) -> dict:
    \"\"\"Mengaktifkan Tab Developer di Ribbon file Excel (.xlsb/.xlsm).\"\"\"
    return optimize_ribbon_developer(filePath)

def main():
    mcp.run()

if __name__ == "__main__":
    main()
