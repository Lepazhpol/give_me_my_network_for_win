# table_merger_full.py — финальный рабочий код с ленивцем и только grid

import os
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Checkbutton, IntVar, Label, Frame, Entry, ttk, PhotoImage
from PIL import Image, ImageTk
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime

SLOTH_IMAGE_PATH = os.path.join(os.path.dirname(__file__), "3e60b959a71738e81e5b5d6f6c5ec03fe9563ab14ba3fc71258be4c84b017d74.png")

YELLOW_FILL = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
RED_FILL = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

selected_acquirers = []
selected_countries = []
df1 = None
df2 = None

def add_sloth_with_text(parent, row=0, column=0, columnspan=3):
    sloth_img_raw = Image.open(SLOTH_IMAGE_PATH).resize((64, 64))
    sloth_img = ImageTk.PhotoImage(sloth_img_raw, master=parent)
    frame = Frame(parent)
    frame.grid(row=row, column=column, columnspan=columnspan, pady=5)
    Label(frame, image=sloth_img).pack(side="left", padx=(10, 5))
    Label(frame, text="Ленивец неспешно пьёт кофе и всё обрабатывает…", font=("Arial", 10, "italic")).pack(side="left")
    frame.image = sloth_img

def get_timestamped_filename():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    return os.path.join(desktop, f"output_{timestamp}.xlsx")

def move_column_after(df, col_to_move, target_col):
    cols = list(df.columns)
    if col_to_move in cols and target_col in cols:
        cols.remove(col_to_move)
        target_index = cols.index(target_col) + 1
        cols.insert(target_index, col_to_move)
        df = df[cols]
    return df

def move_column_first(df, column_name):
    cols = list(df.columns)
    if column_name in cols:
        cols.remove(column_name)
        cols.insert(0, column_name)
        df = df[cols]
    return df

def merge_tables_filtered(progress_window=None):
    global df1, df2, selected_acquirers, selected_countries

    key_col_df1 = 'Код Е100 / E100ID'
    insert_col_df1 = 'S/N Терминала'
    key_col_df2 = 'UID_x0020_АЗС'
    source_col_df2 = 'S_x002F_N_x0020_Терминала'
    acquirer_col_df1 = 'Эквайер / Acquirer'
    country_col_df1 = 'Страна / Country'

    df_filtered = df1[
        (df1[acquirer_col_df1].isin(selected_acquirers)) &
        (df1[country_col_df1].isin(selected_countries))
    ].copy()

    if insert_col_df1 not in df_filtered.columns:
        df_filtered[insert_col_df1] = None

    match_dict = df2.set_index(key_col_df2)[source_col_df2].to_dict()
    duplicates = df_filtered[key_col_df1][df_filtered[key_col_df1].duplicated(keep=False)]

    df_filtered[insert_col_df1] = df_filtered[key_col_df1].map(match_dict)
    df_filtered = move_column_after(df_filtered, insert_col_df1, key_col_df1)

    duplicates_df = df_filtered[df_filtered[key_col_df1].isin(duplicates.values)].copy()
    mismatches_supplier = df_filtered[df_filtered[insert_col_df1].isna()].copy()
    clean_df = df_filtered[
        (~df_filtered[key_col_df1].isin(duplicates.values)) &
        (df_filtered[insert_col_df1].notna())
    ].copy()

    duplicates_df["Источник данных"] = "Supplier"
    mismatches_supplier["Источник данных"] = "Supplier"

    unmatched_webtm_keys = df2[~df2[key_col_df2].isin(df1[key_col_df1])].copy()
    unmatched_webtm_keys = unmatched_webtm_keys[[key_col_df2, source_col_df2]]
    unmatched_webtm_keys.rename(columns={key_col_df2: key_col_df1, source_col_df2: insert_col_df1}, inplace=True)
    unmatched_webtm_keys["Источник данных"] = "WebTM"

    mismatches_df = pd.concat([mismatches_supplier, unmatched_webtm_keys], ignore_index=True)
    duplicates_df = move_column_first(duplicates_df, "Источник данных")
    mismatches_df = move_column_first(mismatches_df, "Источник данных")

    output_file = get_timestamped_filename()

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        clean_df.to_excel(writer, index=False, sheet_name='Все')
        duplicates_df.to_excel(writer, index=False, sheet_name='Задвоения')
        mismatches_df.to_excel(writer, index=False, sheet_name='Несовпадения')

    wb = load_workbook(output_file)
    for sheet, highlight_dups, highlight_miss in [
        ("Задвоения", True, False),
        ("Несовпадения", False, True)
    ]:
        ws = wb[sheet]
        for row in range(2, ws.max_row + 1):
            e100id = ws.cell(row=row, column=3).value
            value = ws.cell(row=row, column=4).value
            if highlight_dups and e100id:
                ws.cell(row=row, column=3).fill = YELLOW_FILL
            if highlight_miss and (value is None or str(value).strip() == ""):
                ws.cell(row=row, column=3).fill = RED_FILL
                ws.cell(row=row, column=4).fill = RED_FILL

    wb.save(output_file)
    if progress_window:
        progress_window.destroy()
    messagebox.showinfo("Готово", f"Файл сохранён на рабочий стол: {os.path.basename(output_file)}")
def show_progress_and_generate(top_window):
    top_window.destroy()
    progress = Toplevel()
    progress.title("Ленивец работает...")
    progress.geometry("400x200")
    progress.resizable(False, False)
    add_sloth_with_text(progress, row=0)
    bar = ttk.Progressbar(progress, mode="indeterminate")
    bar.grid(row=1, column=0, padx=30, pady=20, columnspan=3, sticky="ew")
    bar.start(10)
    progress.update()
    progress.after(500, lambda: merge_tables_filtered(progress))

def open_acquirer_country_window():
    global df1
    top = Toplevel()
    top.title("Выбор эквайеров и стран")
    top.grid_columnconfigure(0, weight=1)
    top.grid_columnconfigure(1, weight=1)
    add_sloth_with_text(top, row=0)

    acquirers_all = sorted(df1['Эквайер / Acquirer'].dropna().unique())
    countries_all = sorted(df1['Страна / Country'].dropna().unique())
    acquirer_vars = {}
    acquirer_checks = {}
    country_vars = {}
    country_checks = {}

    def update_visibility(vars_dict, checks_dict, query):
        for name, chk in checks_dict.items():
            if query in name.lower():
                chk.grid()
            else:
                chk.grid_remove()

    search_acq = Entry(top)
    search_acq.grid(row=1, column=0, padx=10, pady=(10, 0), sticky='ew')
    search_country = Entry(top)
    search_country.grid(row=1, column=1, padx=10, pady=(10, 0), sticky='ew')

    Label(top, text="Эквайеры", font=('Arial', 10, 'bold')).grid(row=2, column=0)
    Label(top, text="Страны", font=('Arial', 10, 'bold')).grid(row=2, column=1)

    frame_acq = Frame(top)
    frame_acq.grid(row=3, column=0, sticky="nsew", padx=10, pady=5)
    canvas_acq = tk.Canvas(frame_acq, height=300, width=250)
    scrollbar_acq = tk.Scrollbar(frame_acq, orient="vertical", command=canvas_acq.yview)
    scrollable_acq = Frame(canvas_acq)
    canvas_acq.create_window((0, 0), window=scrollable_acq, anchor="nw")
    canvas_acq.configure(yscrollcommand=scrollbar_acq.set)
    canvas_acq.pack(side="left", fill="both", expand=True)
    scrollbar_acq.pack(side="right", fill="y")

    for i, val in enumerate(acquirers_all):
        var = IntVar(value=1)
        chk = Checkbutton(scrollable_acq, text=val, variable=var)
        chk.grid(row=i, column=0, sticky="w")
        acquirer_vars[val] = var
        acquirer_checks[val] = chk

    frame_country = Frame(top)
    frame_country.grid(row=3, column=1, sticky="nsew", padx=10, pady=5)
    canvas_country = tk.Canvas(frame_country, height=300, width=250)
    scrollbar_country = tk.Scrollbar(frame_country, orient="vertical", command=canvas_country.yview)
    scrollable_country = Frame(canvas_country)
    canvas_country.create_window((0, 0), window=scrollable_country, anchor="nw")
    canvas_country.configure(yscrollcommand=scrollbar_country.set)
    canvas_country.pack(side="left", fill="both", expand=True)
    scrollbar_country.pack(side="right", fill="y")

    for i, val in enumerate(countries_all):
        var = IntVar(value=1)
        chk = Checkbutton(scrollable_country, text=val, variable=var)
        chk.grid(row=i, column=0, sticky="w")
        country_vars[val] = var
        country_checks[val] = chk

    search_acq.bind("<KeyRelease>", lambda e: update_visibility(acquirer_vars, acquirer_checks, search_acq.get().lower()))
    search_country.bind("<KeyRelease>", lambda e: update_visibility(country_vars, country_checks, search_country.get().lower()))

    def uncheck_all_acquirers():
        for var in acquirer_vars.values():
            var.set(0)

    def uncheck_all_countries():
        for var in country_vars.values():
            var.set(0)

    def generate():
        global selected_acquirers, selected_countries
        selected_acquirers = [name for name, var in acquirer_vars.items() if var.get()]
        selected_countries = [name for name, var in country_vars.items() if var.get()]
        if not selected_acquirers or not selected_countries:
            messagebox.showwarning("Внимание", "Выберите хотя бы одного эквайера и одну страну.")
            return
        show_progress_and_generate(top)

    btns = Frame(top)
    btns.grid(row=4, column=0, columnspan=2, pady=10)
    tk.Button(btns, text="Снять все эквайеры", command=uncheck_all_acquirers).pack(side="left", padx=5)
    tk.Button(btns, text="Снять все страны", command=uncheck_all_countries).pack(side="left", padx=5)
    tk.Button(btns, text="Генерировать", command=generate).pack(side="left", padx=10)

def analyze_files():
    global df1, df2
    file1 = entry1.get()
    file2 = entry2.get()
    if not file1 or not file2:
        messagebox.showwarning("Ошибка", "Выберите оба файла.")
        return
    try:
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)
        if 'Эквайер / Acquirer' not in df1.columns or 'Страна / Country' not in df1.columns:
            messagebox.showerror("Ошибка", "В файле 1 не найдены нужные колонки.")
            return
        open_acquirer_country_window()
    except Exception as e:
        messagebox.showerror("Ошибка анализа", str(e))

def select_file(entry):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

# --- Главное окно ---
root = tk.Tk()
root.title("Объединение таблиц по E100ID")
frame_width = 820
frame_height = 280
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_top = int(screen_height / 2 - frame_height / 2)
position_left = int(screen_width / 2 - frame_width / 2)
root.geometry(f"{frame_width}x{frame_height}+{position_left}+{position_top}")
root.minsize(frame_width, frame_height)

add_sloth_with_text(root, row=0)

tk.Label(root, text="Выгрузка из Supplier:").grid(row=1, column=0, sticky="e", padx=10, pady=10)
entry1 = tk.Entry(root, width=50)
entry1.grid(row=1, column=1, padx=5, pady=10)
tk.Button(root, text="Выбрать...", command=lambda: select_file(entry1)).grid(row=1, column=2, padx=10, pady=10)

tk.Label(root, text="Выгрузка из WebTM:").grid(row=2, column=0, sticky="e", padx=10, pady=10)
entry2 = tk.Entry(root, width=50)
entry2.grid(row=2, column=1, padx=5, pady=10)
tk.Button(root, text="Выбрать...", command=lambda: select_file(entry2)).grid(row=2, column=2, padx=10, pady=10)

tk.Label(root, text="За эту программу можете благодарить одного уставшего Пашку, который не дождался… 😄", fg="gray").grid(row=4, column=0, columnspan=3, pady=(5, 10))

tk.Button(root, text="Анализировать", command=analyze_files, height=2, width=20).grid(row=3, column=1, pady=10)

root.mainloop()
