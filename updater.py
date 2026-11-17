import os
import sys
import zipfile
import urllib.request
import tempfile
import shutil
import pythoncom
import win32com.client as win32

def run_update(target_path_or_name: str, github_zip_url: str):
    print(f"=== MEMULAI UPDATE UNTUK: {target_path_or_name} ===")
    
    # Tentukan path file target secara fleksibel
    if os.path.exists(target_path_or_name):
        target_file = os.path.abspath(target_path_or_name)
    else:
        # Cari di ROOT Documents/excel_files atau folder saat ini
        docs_dir = os.path.join(os.path.expanduser("~"), "Documents", "excel_files")
        candidates = [target_path_or_name, os.path.join(docs_dir, target_path_or_name)]
        if not target_path_or_name.endswith(('.xlsb', '.xlsm')):
            candidates += [target_path_or_name + '.xlsb', os.path.join(docs_dir, target_path_or_name + '.xlsb')]
        
        target_file = None
        for c in candidates:
            if os.path.exists(c):
                target_file = os.path.abspath(c)
                break
        
        if not target_file:
            print(f"[ERROR] File '{target_path_or_name}' tidak ditemukan!")
            sys.exit(1)

    print(f"File target ditemukan: {target_file}")
    file_name = os.path.basename(target_file)

    # 1. Cari dan tutup HANYA workbook target yang sedang terbuka
    try:
        pythoncom.CoInitialize()
        app = win32.Dispatch("Excel.Application")
        found = False
        for wb in app.Workbooks:
            if file_name.lower() in wb.Name.lower():
                print(f"Menutup workbook target: {wb.Name}")
                wb.Close(SaveChanges=False)
                found = True
        if not found:
            print("Workbook target tidak sedang aktif di Excel (lanjut update file fisik).")
    except Exception as e:
        print(f"Catatan COM Excel: {e}")

    # 2. Proses Download & Update via Python (Tanpa 7-Zip!)
    temp_dir = tempfile.mkdtemp()
    backup_path = target_file + '.bak'
    try:
        print("[1/4] Mengunduh paket update dari GitHub...")
        zip_path = os.path.join(temp_dir, "update.zip")
        urllib.request.urlretrieve(github_zip_url, zip_path)
        
        extract_dir = os.path.join(temp_dir, "extracted")
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(extract_dir)

        print("[2/4] Membuat backup file lama...")
        shutil.copy2(target_file, backup_path)

        print("[3/4] Menyuntikkan pembaruan sistem (VBA/Ribbon/XML)...")
        old_xlsb_dir = os.path.join(temp_dir, "old_xlsb")
        with zipfile.ZipFile(target_file, 'r') as old_zf:
            old_zf.extractall(old_xlsb_dir)

        # Cari folder sumber di dalam zip GitHub
        source_update = extract_dir
        for root, dirs, files in os.walk(extract_dir):
            if 'customUI' in dirs or 'xl' in dirs:
                source_update = root
                break

        for root, dirs, files in os.walk(source_update):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, source_update)
                dest_path = os.path.join(old_xlsb_dir, rel_path)
                
                os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                shutil.copy2(full_path, dest_path)

        updated_temp_file = os.path.join(temp_dir, "updated.xlsb")
        with zipfile.ZipFile(updated_temp_file, 'w', zipfile.ZIP_DEFLATED) as new_zf:
            for root, dirs, files in os.walk(old_xlsb_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, old_xlsb_dir)
                    new_zf.write(full_path, arcname)

        shutil.move(updated_temp_file, target_file)
        print("[4/4] Update berhasil diselesaikan!")

    except Exception as e:
        print(f"[ERROR] Gagal melakukan update: {e}")
        if os.path.exists(backup_path):
            shutil.copy2(backup_path, target_file)
            print("Restored dari backup.")
        sys.exit(1)
    finally:
        shutil.rmtree(temp_dir)

if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "MASTER_RBK2026.xlsb"
    url = sys.argv[2] if len(sys.argv) > 2 else "https://codeload.github.com/arv-fazriansyah/updateARB/zip/refs/heads/main"
    run_update(target, url)
