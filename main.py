import tkinter as tk
from tkinter import ttk, messagebox
import csv
from datetime import datetime
import pandas as pd

# Hàm lưu dữ liệu vào file CSV
def save_to_csv():
    data = [
        emp_id.get(),
        emp_name.get(),
        emp_unit.get(),
        emp_title.get(),
        emp_dob.get(),
        emp_gender.get(),
        emp_id_num.get(),
        emp_id_issue_date.get(),
        emp_id_issue_place.get()
    ]
    with open("employee_data.csv", mode="a", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(data)
    messagebox.showinfo("Thông báo", "Lưu dữ liệu thành công!")
    clear_fields()

# Hàm xóa dữ liệu trong các trường nhập
def clear_fields():
    emp_id.set("")
    emp_name.set("")
    emp_unit.set("")
    emp_title.set("")
    emp_dob.set("")
    emp_gender.set("Nam")
    emp_id_num.set("")
    emp_id_issue_date.set("")
    emp_id_issue_place.set("")

# Hàm hiển thị danh sách nhân viên có sinh nhật hôm nay
def show_today_birthdays():
    try:
        today = datetime.today().strftime("%d/%m/%Y")
        with open("employee_data.csv", mode="r", encoding="utf-8") as file:
            reader = csv.reader(file)
            birthdays = [row for row in reader if row[4] == today]
        if birthdays:
            messagebox.showinfo("Sinh nhật hôm nay", "\n".join([f"{row[1]} ({row[0]})" for row in birthdays]))
        else:
            messagebox.showinfo("Sinh nhật hôm nay", "Không có nhân viên nào sinh nhật hôm nay!")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "Không tìm thấy dữ liệu!")

# Hàm xuất danh sách nhân viên ra file Excel
def export_to_excel():
    try:
        df = pd.read_csv("employee_data.csv", header=None, names=["Mã NV", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính", "Số CMND", "Ngày cấp", "Nơi cấp"])
        df["Ngày sinh"] = pd.to_datetime(df["Ngày sinh"], format="%d/%m/%Y", errors="coerce")
        df = df.sort_values(by="Ngày sinh", ascending=False)
        df.to_excel("employee_list.xlsx", index=False)
        messagebox.showinfo("Thông báo", "Xuất danh sách thành công!")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể xuất danh sách: {e}")

# Giao diện chính
root = tk.Tk()
root.title("Quản lý thông tin nhân viên")
root.geometry("600x400")

# Biến lưu trữ thông tin nhân viên
emp_id = tk.StringVar()
emp_name = tk.StringVar()
emp_unit = tk.StringVar()
emp_title = tk.StringVar()
emp_dob = tk.StringVar()
emp_gender = tk.StringVar(value="Nam")
emp_id_num = tk.StringVar()
emp_id_issue_date = tk.StringVar()
emp_id_issue_place = tk.StringVar()

# Layout
lbl = tk.Label(master= root, text="Thông tin nhân viên", font=("Helvetica", 25), anchor="w")
lbl.pack(fill="x",padx=5, pady=5)


frame = ttk.LabelFrame(root)
frame.pack(fill="both", expand="yes", padx=10, pady=10)

ttk.Label(frame, text="Mã NV:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
ttk.Entry(frame, textvariable=emp_id).grid(row=0, column=1, padx=5, pady=5, sticky="w")

ttk.Label(frame, text="Tên:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
ttk.Entry(frame, textvariable=emp_name).grid(row=0, column=3, padx=5, pady=5, sticky="w")

ttk.Label(frame, text="Đơn vị:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
ttk.Entry(frame, textvariable=emp_unit).grid(row=1, column=1, padx=5, pady=5, sticky="w")

ttk.Label(frame, text="Chức danh:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
ttk.Entry(frame, textvariable=emp_title).grid(row=1, column=3, padx=5, pady=5, sticky="w")

ttk.Label(frame, text="Ngày sinh (DD/MM/YYYY):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
ttk.Entry(frame, textvariable=emp_dob).grid(row=2, column=1, padx=5, pady=5, sticky="w")

ttk.Label(frame, text="Giới tính:").grid(row=2, column=2, padx=5, pady=5, sticky="w")
ttk.Radiobutton(frame, text="Nam", variable=emp_gender, value="Nam").grid(row=2, column=3, sticky="w")
ttk.Radiobutton(frame, text="Nữ", variable=emp_gender, value="Nữ").grid(row=2, column=3, padx=60, sticky="w")

ttk.Label(frame, text="Số CMND:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
ttk.Entry(frame, textvariable=emp_id_num).grid(row=3, column=1, padx=5, pady=5, sticky="w")

ttk.Label(frame, text="Ngày cấp (DD/MM/YYYY):").grid(row=3, column=2, padx=5, pady=5, sticky="w")
ttk.Entry(frame, textvariable=emp_id_issue_date).grid(row=3, column=3, padx=5, pady=5, sticky="w")

ttk.Label(frame, text="Nơi cấp:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
ttk.Entry(frame, textvariable=emp_id_issue_place).grid(row=4, column=1, padx=5, pady=5, sticky="w")

# Nút chức năng
btn_frame = ttk.Frame(root)
btn_frame.pack(fill="x", padx=10, pady=10)

ttk.Button(btn_frame, text="Lưu", command=save_to_csv).pack(side="left", padx=5, pady=5)
ttk.Button(btn_frame, text="Sinh nhật hôm nay", command=show_today_birthdays).pack(side="left", padx=5, pady=5)
ttk.Button(btn_frame, text="Xuất danh sách", command=export_to_excel).pack(side="left", padx=5, pady=5)
ttk.Button(btn_frame, text="Xóa thông tin", command=clear_fields).pack(side="left", padx=5, pady=5)

root.mainloop()
