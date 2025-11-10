import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os

# ======= Functions =======

def cap_nhat_sheets(file_path_var, combobox):
    """Cập nhật danh sách sheet khi chọn file"""
    path = file_path_var.get()
    if path:
        try:
            xls = pd.ExcelFile(path)
            combobox['values'] = xls.sheet_names
            if xls.sheet_names:
                combobox.current(0)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể đọc file:\n{e}")

# ---- TT50 ----
def hien_frame_tt50():
    frame_tt80.pack_forget()
    frame_tt50.pack(fill="both", expand=True, padx=5, pady=5)

def xu_ly_tt50():
    try:
        log_text.delete(1.0, tk.END)  # clear log

        sheet_khoa = sheet_khoa_cb.get()
        sheet_tt50 = sheet_tt50_cb.get()
        if not sheet_khoa or not sheet_tt50:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn sheet cho cả 2 file!")
            return

        df_khoa = pd.read_excel(file_khoa_path.get(), sheet_name=sheet_khoa, header=None)
        df_tt50 = pd.read_excel(file_tt50_path.get(), sheet_name=sheet_tt50, header=None)
        df_tt50 = df_tt50[df_tt50[0].notna()]
        df_khoa[5] = ""

        mapping_pt = {2: "PTĐB", 3: "PT1", 4: "PT2", 5: "PT3"}
        mapping_tt = {6: "TTĐB", 7: "TT1", 8: "TT2", 9: "TT3"}

        for idx, row in df_khoa.iterrows():
            if pd.isna(row[0]):
                continue
            ten = str(row[3]).strip().lower()
            if not ten or ten == "nan":
                continue
            log_text.insert(tk.END, f"🔍 Đang kiểm tra (dòng {idx+1}): {row[3]}\n")
            found = df_tt50[df_tt50[1].astype(str).str.strip().str.lower() == ten]
            loai = ""
            if not found.empty:
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
                df_khoa.at[idx, 5] = loai
                log_text.insert(tk.END, f"   ✅ (dòng {idx+1}) Tìm thấy trong TT50 → {loai}\n\n")
            elif found.empty:
                log_text.insert(tk.END, f"   ❌ (dòng {idx+1}) Không tìm thấy trong TT50\n\n")
            else:
                df_khoa.at[idx, 5] = ""
                log_text.insert(tk.END, f"   ⚠️ (dòng {idx+1}) Tìm thấy trong TT50 nhưng không xác định loại PT/TT\n\n")

        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        sheet_name = sheet_khoa_cb.get().replace(" ", "_")
        output_file = os.path.join(desktop, f"Khoa_{sheet_name}_output.xlsx")
        df_khoa.to_excel(output_file, index=False, header=False)
        messagebox.showinfo("Hoàn tất", f"Đã tạo file trên Desktop:\n{output_file}")

    except Exception as e:
        messagebox.showerror("Lỗi", str(e))

# ---- TT80 ----
def hien_frame_tt80():
    frame_tt50.pack_forget()
    frame_tt80.pack(fill="both", expand=True, padx=5, pady=5)

def xu_ly_tt80():
    # Placeholder, logic xử lý tương tự TT50 hoặc tùy chỉnh
    log_text_tt80.delete(1.0, tk.END)
    messagebox.showinfo("TT80", "Chức năng đang được phát triển")

def chon_file(file_path_var, sheet_combobox):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        file_path_var.set(path)
        cap_nhat_sheets(file_path_var, sheet_combobox)

# ======= GUI =======
root = tk.Tk()
root.title("Ứng dụng xử lý PT/TT từ TT50/TT80")
root.geometry("1000x650")
root.configure(bg="#f9f9f9")

# Notebook
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

tab_chucnang = ttk.Frame(notebook)
notebook.add(tab_chucnang, text="Chức năng")

# ----- Menu chức năng con bên trái -----
frame_menu = ttk.Frame(tab_chucnang, width=200)
frame_menu.pack(side="left", fill="y", padx=5, pady=5)

btn_tt50 = ttk.Button(frame_menu, text="Lấy dữ liệu từ TT50", command=hien_frame_tt50)
btn_tt50.pack(fill="x", pady=5)

btn_tt80 = ttk.Button(frame_menu, text="Lấy dữ liệu từ TT80", command=hien_frame_tt80)
btn_tt80.pack(fill="x", pady=5)

# ----- Frame bên phải -----
frame_content = ttk.Frame(tab_chucnang)
frame_content.pack(side="left", fill="both", expand=True, padx=5, pady=5)

# ==== Frame TT50 ====
frame_tt50 = ttk.LabelFrame(frame_content, text="TT50: Import & xử lý dữ liệu", padding=10)

file_khoa_path = tk.StringVar()
file_tt50_path = tk.StringVar()

ttk.Label(frame_tt50, text="File Khoa:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
tk.Entry(frame_tt50, textvariable=file_khoa_path, width=60).grid(row=0, column=1, padx=5, pady=5)
ttk.Button(frame_tt50, text="Chọn", command=lambda: chon_file(file_khoa_path, sheet_khoa_cb)).grid(row=0, column=2, padx=5, pady=5)

ttk.Label(frame_tt50, text="Sheet Khoa:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
sheet_khoa_cb = ttk.Combobox(frame_tt50, state="readonly", width=57)
sheet_khoa_cb.grid(row=1, column=1, padx=5, pady=5)

ttk.Label(frame_tt50, text="File TT50:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
tk.Entry(frame_tt50, textvariable=file_tt50_path, width=60).grid(row=2, column=1, padx=5, pady=5)
ttk.Button(frame_tt50, text="Chọn", command=lambda: chon_file(file_tt50_path, sheet_tt50_cb)).grid(row=2, column=2, padx=5, pady=5)

ttk.Label(frame_tt50, text="Sheet TT50:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
sheet_tt50_cb = ttk.Combobox(frame_tt50, state="readonly", width=57)
sheet_tt50_cb.grid(row=3, column=1, padx=5, pady=5)

btn_xuly = ttk.Button(frame_tt50, text="🚀 Xử lý dữ liệu", command=xu_ly_tt50)
btn_xuly.grid(row=4, column=1, pady=10)

frame_log = ttk.LabelFrame(frame_tt50, text="Kết quả xử lý", padding=10)
frame_log.grid(row=5, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)
frame_tt50.grid_rowconfigure(5, weight=1)
frame_tt50.grid_columnconfigure(1, weight=1)

log_text = tk.Text(frame_log, wrap="word", font=("Consolas", 10))
log_text.pack(side="left", fill="both", expand=True)
scrollbar = ttk.Scrollbar(frame_log, command=log_text.yview)
scrollbar.pack(side="right", fill="y")
log_text.config(yscrollcommand=scrollbar.set)

# ==== Frame TT80 ====
frame_tt80 = ttk.LabelFrame(frame_content, text="TT80: Import & xử lý dữ liệu", padding=10)
log_text_tt80 = tk.Text(frame_tt80, wrap="word", font=("Consolas", 10))
log_text_tt80.pack(side="left", fill="both", expand=True)
scrollbar_tt80 = ttk.Scrollbar(frame_tt80, command=log_text_tt80.yview)
scrollbar_tt80.pack(side="right", fill="y")
log_text_tt80.config(yscrollcommand=scrollbar_tt80.set)

btn_xuly_tt80 = ttk.Button(frame_tt80, text="🚀 Xử lý dữ liệu TT80", command=xu_ly_tt80)
btn_xuly_tt80.pack(pady=10)

root.mainloop()
