import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os

def cap_nhat_sheets(file_path_var, combobox):
    """Cập nhật danh sách sheet khi chọn file"""
    path = file_path_var.get()
    if path:
        try:
            xls = pd.ExcelFile(path)
            combobox['values'] = xls.sheet_names
            if xls.sheet_names:
                combobox.current(0)  # chọn sheet đầu tiên mặc định
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể đọc file:\n{e}")

def xu_ly():
    try:
        log_text.delete(1.0, tk.END)  # clear log

        # Lấy sheet được chọn
        sheet_khoa = sheet_khoa_cb.get()
        sheet_tt50 = sheet_tt50_cb.get()

        if not sheet_khoa or not sheet_tt50:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn sheet cho cả 2 file!")
            return

        # Đọc sheet được chọn
        df_khoa = pd.read_excel(file_khoa_path.get(), sheet_name=sheet_khoa, header=None)
        df_tt50 = pd.read_excel(file_tt50_path.get(), sheet_name=sheet_tt50, header=None)

        # Loại bỏ các dòng TT50 không có STT (cột A = index 0)
        df_tt50 = df_tt50[df_tt50[0].notna()]

        # Ánh xạ loại
        mapping_pt = {2: "PTĐB", 3: "PT1", 4: "PT2", 5: "PT3"}
        mapping_tt = {6: "TTĐB", 7: "TT1", 8: "TT2", 9: "TT3"}

        # Lặp qua từng đầu mục trong file Khoa (cột D = index 3)
        for idx, row in df_khoa.iterrows():
            if pd.isna(row[0]):  # bỏ qua dòng không có STT
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
                    df_khoa.at[idx, 5] = loai  # ghi vào cột F
                    log_text.insert(tk.END, f"   ✅ (dòng {idx+1}) Tìm thấy trong TT50 → {loai}\n\n")
                else:
                    log_text.insert(tk.END, f"   ⚠️ (dòng {idx+1}) Tìm thấy trong TT50 nhưng không xác định loại PT/TT\n\n")
            else:
                # Nếu không tìm thấy loại PT/TT → để trống
                df_khoa.at[idx, 5] = ""
                log_text.insert(tk.END, f"   ❌ (dòng {idx+1}) Không tìm thấy trong TT50\n\n")

        # Xuất file ra Desktop với tên dựa trên sheet
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        sheet_name = sheet_khoa_cb.get().replace(" ", "_")  # thay khoảng trắng bằng _
        output_file = os.path.join(desktop, f"Khoa_{sheet_name}_output.xlsx")
        df_khoa.to_excel(output_file, index=False, header=False)
        messagebox.showinfo("Hoàn tất", f"Đã tạo file trên Desktop:\n{output_file}")

    except Exception as e:
        messagebox.showerror("Lỗi", str(e))

def chon_file(file_path_var, sheet_combobox):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        file_path_var.set(path)
        cap_nhat_sheets(file_path_var, sheet_combobox)

# GUI
root = tk.Tk()
root.title("Lấy loại PT/TT từ TT50 (sheet tùy chọn)")
root.geometry("800x550")
root.configure(bg="#f9f9f9")

frame_files = ttk.LabelFrame(root, text="Chọn file Excel", padding=10)
frame_files.pack(fill="x", padx=15, pady=10)

file_khoa_path = tk.StringVar()
file_tt50_path = tk.StringVar()

# File Khoa
ttk.Label(frame_files, text="File Khoa:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
tk.Entry(frame_files, textvariable=file_khoa_path, width=60).grid(row=0, column=1, padx=5, pady=5)
ttk.Button(frame_files, text="Chọn", command=lambda: chon_file(file_khoa_path, sheet_khoa_cb)).grid(row=0, column=2, padx=5, pady=5)

# Sheet Khoa
ttk.Label(frame_files, text="Sheet Khoa:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
sheet_khoa_cb = ttk.Combobox(frame_files, state="readonly", width=57)
sheet_khoa_cb.grid(row=1, column=1, padx=5, pady=5)

# File TT50
ttk.Label(frame_files, text="File TT50:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
tk.Entry(frame_files, textvariable=file_tt50_path, width=60).grid(row=2, column=1, padx=5, pady=5)
ttk.Button(frame_files, text="Chọn", command=lambda: chon_file(file_tt50_path, sheet_tt50_cb)).grid(row=2, column=2, padx=5, pady=5)

# Sheet TT50
ttk.Label(frame_files, text="Sheet TT50:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
sheet_tt50_cb = ttk.Combobox(frame_files, state="readonly", width=57)
sheet_tt50_cb.grid(row=3, column=1, padx=5, pady=5)

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
