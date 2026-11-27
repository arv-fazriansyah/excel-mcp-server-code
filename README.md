# Excel MCP Power & Engine Guide

Panduan operasional cepat untuk menggunakan engine Excel berbasis COM (`pywin32`) di folder ini.

## Quick Install (Plug & Play)

Satu perintah (PowerShell) — clone, buat venv, install dependency, smoke test, daftar ke OpenCode:

```powershell
irm https://raw.githubusercontent.com/arv-fazriansyah/excel-mcp-server-code/main/install.ps1 | iex
```

Atau jalankan dari folder repo:

```powershell
.\install.ps1
```

Opsi:
- `-InstallDir "D:\tools\excel-mcp"` — lokasi instalasi kustom
- `-UseSystemPython` — pakai Python sistem (tanpa venv)
- `-SkipSmokeTest` — lewati smoke test

Prasyarat: Windows, Python 3.10+, Git, Microsoft Excel terpasang, OpenCode CLI di PATH.

Setelah instalasi selesai, restart OpenCode (atau `opencode service restart`) agar MCP `excel` dimuat.

## Lokasi & Struktur Proyek
- **`excel_engine.py`**: Core engine tunggal (Manajemen session, Sheet, Range, Formula, Table, Named Range, VBA).
- **`server.py`**: Entry point FastMCP server.
- **`tests/smoke_test.py`**: Suite pengujian otomatis.

## Aturan Utama
1. **Windows + pywin32**: Wajib berjalan di Windows dengan Microsoft Excel terpasang.
2. **Tanpa Proses Zombie**: Setiap sesi workbook harus ditutup (`session_close`) agar tidak ada proses `EXCEL.EXE` yang menggantung di Task Manager.
3. **Format Didukung**: `.xlsx`, `.xlsm`, `.xlsb`, `.xls`.
4. **VBA / Macro**: Memerlukan pengaturan Excel *"Trust access to the VBA project object model"* aktif.

## Contoh Penggunaan Cepat via Python (CLI / Script)

Path file secara otomatis mendeteksi folder `EXCEL_MCP_ROOT` (default: `~/Documents/excel_files`) atau path absolut.

```python
import os
from excel_engine import ROOT, session_open, session_info, range_get_values, session_close

# 1. Buka workbook contoh secara dinamis dari root folder
file_path = os.path.join(ROOT, "contoh_workbook.xlsb")
res = session_open(file_path, read_only=True)
sid = res["sessionId"]

# 2. Baca info / struktur sheet
info = session_info(sid)
print("Sheets:", [s["name"] for s in info["sheets"]])

# 3. Baca data range sel
data = range_get_values(sid, "Sheet1", "A1:E10")
print(data["values"])

# 4. Tutup sesi dengan bersih
session_close(sid)
```

## Menjalankan Smoke Test
```powershell
python tests/smoke_test.py
```
