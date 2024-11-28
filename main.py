import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox

# رابط ملف الإكسل المباشر
EXCEL_FILE_URL = "https://dl.dropbox.com/scl/fi/7dbconhwt7hkwxzgroats/st.xlsx?rlkey=dwg7ke4curmpujob5g0x9a72l&st=dt07s2dd&dl=0"

def read_excel_from_url():
    """قراءة ملف الإكسل مباشرة من الرابط."""
    try:
        df = pd.read_excel(EXCEL_FILE_URL)
        print("تم تحميل الملف بنجاح!")
        return df
    except Exception as e:
        print(f"حدث خطأ أثناء تحميل الملف: {e}")
        return None

def search_in_excel(df, column_name, search_value):
    """بحث عن النتائج في ملف الإكسل بتطابق تام."""
    results = df[df[column_name].astype(str) == search_value]
    return results

def on_search():
    """تنفيذ البحث عند الضغط على الزر."""
    search_value = search_entry.get().strip()
    if not search_value:
        messagebox.showwarning("تنبيه", "يرجى إدخال قيمة للبحث")
        return
    
    column_name = "اسم الطالب"
    results = search_in_excel(df, column_name, search_value)
    
    # تنظيف النتائج القديمة
    for widget in results_frame.winfo_children():
        widget.destroy()
    
    if results.empty:
        messagebox.showinfo("نتائج البحث", "لا توجد نتائج مطابقة")
    else:
        for index, row in results.iterrows():
            for col_name, value in row.items():
                row_frame = tk.Frame(results_frame, bg="#f9f9f9", relief="groove", borderwidth=1)
                row_frame.pack(fill="x", pady=2, padx=10)

                tk.Label(
                    row_frame,
                    text=f"{col_name}:",
                    font=("Arial", 14, "bold"),
                    bg="#4CAF50",
                    fg="white",
                    width=20,
                    anchor="e"  # الاتجاه من اليمين
                ).pack(side="right", padx=5, pady=5)
                
                tk.Label(
                    row_frame,
                    text=str(value),
                    font=("Arial", 14),
                    bg="#ffffff",
                    fg="black",
                    anchor="w"
                ).pack(side="right", padx=5, pady=5)

# قراءة ملف الإكسل من الرابط
df = read_excel_from_url()
if df is None:
    print("لا يمكن تحميل البيانات. تأكد من صحة الرابط.")
    exit()

# إنشاء واجهة المستخدم
root = tk.Tk()
root.title("البحث عن نتائج الطلاب")
root.geometry("900x600")
root.configure(bg="#f2f9ff")  # خلفية التطبيق بلون أزرق فاتح

# عنوان التطبيق
tk.Label(
    root,
    text="البحث عن نتائج الطلاب",
    font=("Arial", 18, "bold"),
    bg="#4CAF50",
    fg="white"
).pack(fill="x", pady=10)

# مربع الإدخال
search_frame = tk.Frame(root, bg="#eaf3ff", relief="raised", borderwidth=2)
search_frame.pack(pady=10, fill="x", padx=20)

tk.Label(
    search_frame,
    text="أدخل اسم الطالب:",
    font=("Arial", 15),
    bg="#eaf3ff"
).pack(side="right", padx=10)

search_entry = tk.Entry(
    search_frame,
    width=30,
    font=("Arial", 15),
    justify="right",
    relief="solid",
    borderwidth=1
)
search_entry.pack(side="right", padx=10)

# زر البحث
search_button = tk.Button(
    root,
    text="بحث",
    command=on_search,
    font=("Arial", 15, "bold"),
    bg="#007BFF",
    fg="white",
    relief="raised"
)
search_button.pack(pady=10)

# إطار النتائج
results_frame = tk.Frame(root, bg="#ffffff", relief="groove", borderwidth=2)
results_frame.pack(pady=10, fill="both", expand=True, padx=20)

# تشغيل التطبيق
root.mainloop()
