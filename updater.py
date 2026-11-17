import os
import sys
import zipfile
import urllib.request
import tempfile
import shutil
import pythoncom
import win32com.client as win32

def find_target_excel(arg_input: str = None) -> str:
    """Mencari file target Excel secara cerdas berdasarkan argumen, drag-and-drop, atau smart scan dengan menu pilihan user."""
    # 1. Jika path diberikan langsung (dari VBA atau Drag & Drop), langsung gunakan
    if arg_input and os.path.exists(arg_input) and arg_input.endswith(('.xlsb', '.xlsm')):
        return os.path.abspath(arg_input)
        
    # 2. Smart Scan: Cari di folder updater, Documents\excel_files, Documents, dan Downloads
    search_dirs = [
        os.path.dirname(os.path.abspath(__file__)),
        os.path.join(os.path.expanduser("~"), "Documents", "excel_files"),
        os.path.join(os.path.expanduser("~"), "Documents"),
        os.path.join(os.path.expanduser("~"), "Downloads")
    ]
    
    found_files = []
    for d in search_dirs:
        if os.path.exists(d):
            for root, dirs, files in os.walk(d):
                if "backup" in root.lower() or "temp" in root.lower():
                    continue
                if root.count(os.sep) - d.count(os.sep) > 2: 
                    continue
                for f in files:
                    if "rbk" in f.lower() or "master" in f.lower() or "arb" in f.lower():
                        if f.endswith(('.xlsb', '.xlsm')):
                            found_files.append(os.path.join(root, f))
                            
    # Hilangkan duplikat dan file backup (.bak atau folder backup)
    found_files = [f for f in list(set(found_files)) if "BACKUP" not in f and not f.endswith('.bak')]
    
    if not found_files:
        print("[ERROR] Tidak ditemukan file aplikasi Excel (.xlsb/.xlsm) di komputer ini!")
        print("Tip: Seret (Drag) file Excel Anda dan lepaskan di atas ikon updater.exe")
        return None
        
    # Selalu tampilkan daftar file agar user dapat memilih secara interaktif
    print("\n" + "=" * 60)
    print(" DAFTAR FILE APLIKASI EXCEL YANG DITEMUKAN:")
    print("=" * 60)
    for idx, ff in enumerate(found_files):
        print(f"  [{idx + 1}] {ff}")
    print("=" * 60)
    
    while True:
        choice = input(f"Pilih nomor file yang ingin di-update [1-{len(found_files)}] (atau tekan Ctrl+C untuk batal): ").strip()
        try:
            sel_idx = int(choice) - 1
            if 0 <= sel_idx < len(found_files):
                selected = found_files[sel_idx]
                print(f"\n[OK] Anda memilih: {selected}\n")
                return selected
            else:
                print("[ERROR] Nomor pilihan tidak valid. Silakan coba lagi.")
        except ValueError:
            print("[ERROR] Masukkan angka yang valid.")

def run_update(github_zip_url: str, target_path_input: str, version: str):
    print("=== PROGRAM PEMBARUAN APLIKASI (UPDATER) ===")
    
    target_file = find_target_excel(target_path_input)
    if not target_file:
        sys.exit(1)
        
    print(f"Target terpilih: {target_file}")
    target_dir = os.path.dirname(target_file)
    file_name = os.path.basename(target_file)
    backup_dir = os.path.join(target_dir, "BACKUP APLIKASI ARB 2026")
    
    temp_dir = tempfile.mkdtemp()
    backup_path = target_file + '.bak'
    
    try:
        print("\n[1/10] Menyiapkan update manager...")
        print("[2/10] Membersihkan file sementara...")
        
        print("[3/10] Mengecek koneksi internet...")
        urllib.request.urlopen("https://google.com", timeout=5)
        
        print(f"[5/10] Menutup file {file_name} (File Excel lain aman)...")
        try:
            pythoncom.CoInitialize()
            app = win32.Dispatch("Excel.Application")
            for wb in app.Workbooks:
                if os.path.abspath(wb.FullName).lower() == target_file.lower():
                    wb.Close(SaveChanges=False)
                    break
        except Exception:
            pass
            
        print("[6/10] Membuat backup file...")
        os.makedirs(backup_dir, exist_ok=True)
        shutil.copy2(target_file, backup_path)
        
        print("[7/10] Mengunduh paket update dari GitHub...")
        zip_path = os.path.join(temp_dir, "update.zip")
        urllib.request.urlretrieve(github_zip_url, zip_path)
        
        print("[8/10] Mengekstrak pembaruan...")
        extract_dir = os.path.join(temp_dir, "extracted")
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(extract_dir)
            
        source_update = extract_dir
        for root, dirs, files in os.walk(extract_dir):
            if 'customUI' in dirs or 'xl' in dirs:
                source_update = root
                break
                
        print("[9/10] Menyuntikkan pembaruan sistem ke file Excel...")
        old_xlsb_dir = os.path.join(temp_dir, "old_xlsb")
        with zipfile.ZipFile(target_file, 'r') as old_zf:
            old_zf.extractall(old_xlsb_dir)
            
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
        
        print("[10/10] Finishing & Clean up...")
        base_name_only = file_name
        for prefix in ["v24.", "v25.", "v26."]:
            if prefix in base_name_only:
                parts = base_name_only.split("_", 1)
                if len(parts) > 1: base_name_only = parts[1]
                break
                
        new_name = f"{version}_{base_name_only}"
        new_path = os.path.join(target_dir, new_name)
        if target_file != new_path:
            if os.path.exists(new_path): os.remove(new_path)
            os.rename(target_file, new_path)
            print(f"  File direname menjadi: {new_name}")
            
        print("=" * 50)
        print("[OK] PROSES UPDATE SELESAI DENGAN SUKSES!")
        print(f"File diperbarui menjadi: {new_name}")
        print("=" * 50)
        
        input("\nTekan Enter untuk keluar...")
        
    except Exception as e:
        print(f"\n[ERROR] Update gagal: {e}")
        if os.path.exists(backup_path):
            shutil.copy2(backup_path, target_file)
            print("Dipulihkan kembali dari file backup.")
        input("\nTekan Enter untuk keluar...")
        sys.exit(1)
    finally:
        shutil.rmtree(temp_dir)

if __name__ == "__main__":
    url = "https://codeload.github.com/arv-fazriansyah/updateARB/zip/refs/heads/main"
    ver = sys.argv[2] if len(sys.argv) > 2 else "v25.11.2025"
    
    # Jika dijalankan manual (klik 2x), arg_path kosong. Jika dari VBA / drag-drop, ambil dari argumen 1.
    arg_path = sys.argv[1] if len(sys.argv) > 1 and os.path.exists(sys.argv[1]) else None
    
    run_update(url, arg_path, ver)
