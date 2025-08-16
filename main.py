import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os

def xu_ly():
    try:
        log_text.delete(1.0, tk.END)  # clear log

        # Đọc file
        df_khoa = pd.read_excel(file_khoa_path.get(), header=None)
        df_tt50 = pd.read_excel(file_tt50_path.get(), header=None)

        # Loại bỏ các dòng TT50 không có STT (cột A = index 0)
        df_tt50 = df_tt50[df_tt50[0].notna()]

        # Ánh xạ loại
        mapping_pt = {2: "PTĐB", 3: "PT1", 4: "PT2", 5: "PT3"}
        mapping_tt = {6: "TTĐB", 7: "TT1", 8: "TT2", 9: "TT3"}

        # Lặp qua từng đầu mục trong file Khoa (cột D = index 3)
        for idx, row in df_khoa.iterrows():
            # Bỏ qua dòng Khoa không có STT (cột A = index 0)
            if pd.isna(row[0]):
                continue

            ten = str(row[3]).strip().lower()
            if not ten or ten == "nan":
                continue

            log_text.insert(tk.END, f"🔍 Đang kiểm tra (dòng {idx+1}): {row[3]}\n")

            # Tìm trong file TT50 (cột B = index 1)
            found = df_tt50[df_tt50[1].astype(str).str.strip().str.lower() == ten]

            if not found.empty:
                loai = None
                for _, r in found.iterrows():
                    for col, label in mapping_pt.items():
                        if str(r[col]).strip().lower() == "x":
                            loai = label
                            break
                    for col, label in mapping_tt.items():
                        if str(r[col]).strip().lower() == "x":
                            loai = label
                            break
                if loai:
                    df_khoa.at[idx, 5] = loai  # ghi vào cột F (index=5)
                    log_text.insert(tk.END, f"   ✅ (dòng {idx+1}) Tìm thấy trong TT50 → {loai}\n\n")
                else:
                    log_text.insert(tk.END, f"   ⚠️ (dòng {idx+1}) Tìm thấy trong TT50 nhưng không xác định loại PT/TT\n\n")
            else:
                log_text.insert(tk.END, f"   ❌ (dòng {idx+1}) Không tìm thấy trong TT50\n\n")

        # Xuất file ra Desktop
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        output_file = os.path.join(desktop, "Khoa_output.xlsx")
        df_khoa.to_excel(output_file, index=False, header=False)
        messagebox.showinfo("Hoàn tất", f"Đã tạo file trên Desktop:\n{output_file}")

    except Exception as e:
        messagebox.showerror("Lỗi", str(e))

def chon_file(entry):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        entry.set(path)

# GUI
root = tk.Tk()
root.title("Lấy ra loại PT/TT từ Tên kỹ thuật trong TT50")
root.geometry("750x500")
root.configure(bg="#f9f9f9")

# Frame chọn file
frame_files = ttk.LabelFrame(root, text="Chọn file Excel", padding=10)
frame_files.pack(fill="x", padx=15, pady=10)

file_khoa_path = tk.StringVar()
file_tt50_path = tk.StringVar()

ttk.Label(frame_files, text="File Khoa:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
ttk.Entry(frame_files, textvariable=file_khoa_path, width=60).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
ttk.Button(frame_files, text="Chọn", command=lambda: chon_file(file_khoa_path)).grid(row=0, column=2, padx=5, pady=5)

ttk.Label(frame_files, text="File TT50:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
ttk.Entry(frame_files, textvariable=file_tt50_path, width=60).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
ttk.Button(frame_files, text="Chọn", command=lambda: chon_file(file_tt50_path)).grid(row=1, column=2, padx=5, pady=5)

# Nút xử lý
btn_xuly = ttk.Button(root, text="🚀 Xử lý dữ liệu", command=xu_ly)
btn_xuly.pack(pady=10)

# Khung log
frame_log = ttk.LabelFrame(root, text="Kết quả xử lý", padding=10)
frame_log.pack(fill="both", expand=True, padx=15, pady=10)

log_text = tk.Text(frame_log, wrap="word", font=("Consolas", 10))
log_text.pack(side="left", fill="both", expand=True)

scrollbar = ttk.Scrollbar(frame_log, command=log_text.yview)
scrollbar.pack(side="right", fill="y")
log_text.config(yscrollcommand=scrollbar.set)

root.mainloop()
