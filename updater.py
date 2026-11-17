import os
import sys
import zipfile
import urllib.request
import tempfile
import shutil
import pythoncom
import win32com.client as win32

def run_update(version: str, target_name: str, github_zip_url: str):
    print(f"=== MEMULAI UPDATE UNTUK: {target_name} ===")
    print(f"Versi: {version}")
    print(f"File: {target_name}.xlsb")
    print("=" * 70)
    
    # --- [1] Menyiapkan update manager ---
    print("[1/10] Menyiapkan update manager...")
    
    # Cari file target di folder Documents/excel_files atau folder saat ini
    docs_dir = os.path.join(os.path.expanduser("~"), "Documents", "excel_files")
    target_file = None
    for search_dir in [os.path.dirname(os.path.abspath(__file__)), docs_dir]:
        if not os.path.exists(search_dir):
            continue
        for f in os.listdir(search_dir):
            if target_name.lower() in f.lower() and f.endswith(('.xlsb', '.xlsm')):
                target_file = os.path.join(search_dir, f)
                break
        if target_file:
            break
    
    if not target_file or not os.path.exists(target_file):
        print(f"[ERROR] FILE TIDAK DITEMUKAN: {target_name}")
        sys.exit(1)
    
    print(f"[OK] PATH: {target_file}")
    target_dir = os.path.dirname(target_file)
    backup_dir = os.path.join(target_dir, "BACKUP APLIKASI ARB 2026")
    
    # --- [2] Membersihkan file sementara ---
    print("[2/10] Membersihkan file sementara...")
    temp_dir = tempfile.mkdtemp()
    try:
        # --- [3] Mengecek koneksi internet ---
        print("[3/10] Mengecek koneksi internet...")
        try:
            urllib.request.urlopen("https://google.com", timeout=5)
        except Exception:
            print("[ERROR] TIDAK ADA KONEKSI INTERNET.")
            sys.exit(1)
        
        # --- [4] Mengecek file target ---
        print("[4/10] Mengecek file ARB...")
        if not os.path.exists(target_file):
            print("[ERROR] FILE TIDAK DITEMUKAN.")
            sys.exit(1)
        
        # --- [5] Menutup file Excel target saja (tidak menutup semua Excel!) ---
        print("[5/10] Menutup file ARB (hanya target, Excel lain aman)...")
        try:
            pythoncom.CoInitialize()
            app = win32.Dispatch("Excel.Application")
            file_name = os.path.basename(target_file)
            for wb in app.Workbooks:
                if file_name.lower() in wb.Name.lower():
                    print(f"  Menutup: {wb.Name}")
                    wb.Close(SaveChanges=False)
        except Exception as e:
            print(f"  Catatan COM: {e}")
        
        # --- [6] Membuat backup ---
        print("[6/10] Membuat backup file ARB...")
        os.makedirs(backup_dir, exist_ok=True)
        shutil.copy2(target_file, backup_dir)
        
        # --- [7] Mengunduh file update dari GitHub ---
        print("[7/10] Mengunduh file update...")
        zip_path = os.path.join(temp_dir, "update.zip")
        urllib.request.urlretrieve(github_zip_url, zip_path)
        
        # --- [8] Mengekstrak file update ---
        print("[8/10] Mengekstrak file update...")
        extract_dir = os.path.join(temp_dir, "extracted")
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(extract_dir)
        
        # Cari folder sumber update (yang berisi customUI/xl)
        source_update = extract_dir
        for root, dirs, files in os.walk(extract_dir):
            if 'customUI' in dirs or 'xl' in dirs:
                source_update = root
                break
        
        # --- [9] Memperbarui file (merge update ke .xlsb) ---
        print("[9/10] Memperbarui file ARB...")
        old_xlsb_dir = os.path.join(temp_dir, "old_xlsb")
        with zipfile.ZipFile(target_file, 'r') as old_zf:
            old_zf.extractall(old_xlsb_dir)
        
        # Timpa file internal VBA/Ribbon/XML dari update baru
        for root, dirs, files in os.walk(source_update):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, source_update)
                dest_path = os.path.join(old_xlsb_dir, rel_path)
                os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                shutil.copy2(full_path, dest_path)
        
        # Bungkus kembali
        updated_temp_file = os.path.join(temp_dir, "updated.xlsb")
        with zipfile.ZipFile(updated_temp_file, 'w', zipfile.ZIP_DEFLATED) as new_zf:
            for root, dirs, files in os.walk(old_xlsb_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, old_xlsb_dir)
                    new_zf.write(full_path, arcname)
        
        shutil.move(updated_temp_file, target_file)
        
        # --- [10] Ganti nama file dengan prefix versi ---
        print("[10/10] Cleaning & Finishing...")
        basename = os.path.basename(target_file)
        # Jika nama file belum memiliki prefix versi, tambahkan
        if not basename.startswith(version):
            new_name = f"{version}_{basename}"
            new_path = os.path.join(target_dir, new_name)
            os.rename(target_file, new_path)
            target_file = new_path
            print(f"  File direname: {new_name}")
        
        print("=" * 70)
        print("[OK] PROSES UPDATE SELESAI.")
        print(f"File updated: {target_file}")
        
    except Exception as e:
        print(f"[ERROR] Gagal melakukan update: {e}")
        if os.path.exists(backup_path):
            shutil.copy2(backup_path, target_file)
            print("Restored dari backup.")
        sys.exit(1)
    finally:
        shutil.rmtree(temp_dir)

if __name__ == "__main__":
    # Argumen: version target_name github_zip_url
    version = sys.argv[1] if len(sys.argv) > 1 else "v25.11.2025"
    target = sys.argv[2] if len(sys.argv) > 2 else "MASTER_RBK2026"
    url = sys.argv[3] if len(sys.argv) > 3 else "https://codeload.github.com/arv-fazriansyah/updateARB/zip/refs/heads/main"
    run_update(version, target, url)
