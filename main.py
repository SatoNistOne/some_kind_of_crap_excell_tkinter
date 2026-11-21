import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

def generate_excel():
    data = {k: v.get().strip() for k, v in entries.items()}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Накладная"
    bold = Font(bold=True)
    title_font = Font(bold=True, size=14)
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws["A1"] = f"ТРАНСПОРТНАЯ НАКЛАДНАЯ № {data['num']} от {data['date']}"
    ws["A1"].font = title_font
    ws.merge_cells("A1:B1")
    rows = [
        ("Грузоотправитель (Наименование, ИНН, адрес):", data["sender"]),
        ("Грузополучатель (Наименование, ИНН, адрес):", data["receiver"]),
        ("Перевозчик (Наименование, ИНН):", data["carrier"]),
        ("Адрес погрузки:", data["load_addr"]),
        ("Адрес выгрузки:", data["unload_addr"]),
        ("Груз (наименование, кол-во мест, масса кг):", data["cargo"]),
        ("Автомобиль (марка, госномер):", data["car"]),
        ("Водитель (ФИО, телефон):", data["driver"]),
        ("Стоимость перевозки (руб.):", data["price"]),
    ]
    for idx, (label, value) in enumerate(rows, start=3):
        c1 = ws.cell(row=idx, column=1, value=label)
        c2 = ws.cell(row=idx, column=2, value=value)
        c1.font = bold
        c1.border = c2.border = border
        c1.alignment = c2.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[idx].height = 28
    sign_row = len(rows) + 5
    ws.cell(row=sign_row, column=1, value="Груз сдал (Отправитель): _______________")
    ws.cell(row=sign_row, column=2, value="Груз принял (Водитель): _______________")
    ws.cell(row=sign_row + 2, column=1, value="Груз сдал (Водитель): __________________")
    ws.cell(row=sign_row + 2, column=2, value="Груз принял (Получатель): _____________")
    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 50
    filename = f"Накладная_{data['num'] or 'bez_nomera'}.xlsx"
    wb.save(filename)
    messagebox.showinfo("Готово", f"Файл сохранен:\n{filename}")

root = tk.Tk()
root.title("Мини-накладная")
root.geometry("720x720")
root.configure(bg="#111827")

fields = [
    ("num", "Номер накладной:", ""),
    ("date", "Дата:", datetime.now().strftime("%d.%m.%Y")),
    ("sender", "Грузоотправитель:", ""),
    ("receiver", "Грузополучатель:", ""),
    ("carrier", "Перевозчик:", ""),
    ("load_addr", "Адрес погрузки:", ""),
    ("unload_addr", "Адрес выгрузки:", ""),
    ("cargo", "Груз (название, места, вес):", ""),
    ("car", "Автомобиль и госномер:", ""),
    ("driver", "Водитель (ФИО, тел.):", ""),
    ("price", "Стоимость перевозки:", ""),
]

entries = {}
for i, (key, label_text, default_val) in enumerate(fields):
    tk.Label(
        root, text=label_text, anchor="w", bg="#111827", fg="#A2B4D8", font=("Segoe UI", 11)
    ).grid(row=i, column=0, padx=(30, 15), pady=8, sticky="w")
    ent = tk.Entry(
        root, width=42, font=("Segoe UI", 11), bg="#1F2937", fg="#F9FAFB", insertbackground="#F9FAFB",
        relief="flat", highlightthickness=1, highlightbackground="#374151", highlightcolor="#3B82F6"
    )
    ent.insert(0, default_val)
    ent.grid(row=i, column=1, padx=(15, 30), pady=8, ipady=6)
    entries[key] = ent

tk.Button(
    root, text="Сформировать Excel", command=generate_excel,
    bg="#3B82F6", fg="#FFFFFF", activebackground="#2563EB", activeforeground="#FFFFFF",
    font=("Segoe UI", 11, "bold"), relief="flat", cursor="hand2", padx=32, pady=10
).grid(row=len(fields), column=0, columnspan=2, pady=24)

root.mainloop()
## сделал рабочий минимум по тз
## ни проверок ни чего то подобного но работает