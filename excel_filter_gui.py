import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from constant import ALL_COLUMNS, EXCEL_TYPES
from excel_processor import excel_process
from tkinter import messagebox

root = tk.Tk()
root.title("Анализ данных")
root.geometry("1000x500")
input_path = tk.StringVar()
output_path = tk.StringVar()
filter_column = tk.StringVar()
filter_value = tk.StringVar()


def choose_input_file() -> None:
    filename = filedialog.askopenfilename(
        title="Выберите файл Excel",
        filetypes=EXCEL_TYPES,
    )
    if filename:
        input_path.set(filename)


def choose_output_file() -> None:
    filename = filedialog.asksaveasfilename(
        title="Сохранить результат как",
        defaultextension=".xlsx",
        filetypes=EXCEL_TYPES,
    )
    if filename:
        output_path.set(filename)


def run_filter(filter_column: str, filter_value: str, output_path: str, input_path: str) -> None:
    try:
        excel_process(filter_column, filter_value, output_path, input_path)
    except ValueError as e:
        messagebox.showerror("Ошибка", str(e))


input_frame = ttk.LabelFrame(root, text="Входной файл", padding=10)
input_frame.pack(fill="x", padx=10, pady=5)
ttk.Entry(input_frame, textvariable=input_path, width=80).pack(side="left", padx=(0, 5))
btn_input = ttk.Button(input_frame, text="Выбор файла для анализа", command=choose_input_file)
btn_input.pack(side="left")


output_frame = ttk.LabelFrame(root, text="Выходной файл", padding=10)
output_frame.pack(fill="x", padx=10, pady=5)
ttk.Entry(output_frame, textvariable=output_path, width=80).pack(side="left", padx=(0, 5))
btn_output = ttk.Button(output_frame, text="Выбор пути для сохранения данных", command=choose_output_file)
btn_output.pack(side="left")


filter_frame = ttk.LabelFrame(root, text="Параметры фильтрации", padding=10)
filter_frame.pack(fill="x", padx=10, pady=5)
ttk.Label(filter_frame, text="Столбец для фильтрации:").grid(row=0, column=0, sticky="w", pady=5)
ttk.Combobox(
    filter_frame,
    textvariable=filter_column,
    values=ALL_COLUMNS,
    state="readonly",
    width=30,
).grid(row=0, column=1, padx=5, pady=5)

ttk.Label(filter_frame, text="Значение для фильтрации:").grid(
    row=1, column=0, sticky="w", pady=5
)
ttk.Entry(filter_frame, textvariable=filter_value, width=30).grid(
    row=1, column=1, padx=5, pady=5
)


btn_filter = ttk.Button(
    root,
    text="Запустить фильтрацию",
    command=lambda: run_filter(
        filter_column.get(),
        filter_value.get(),
        output_path.get(),
        input_path.get(),
    ),
    style="Accent.TButton",
)
btn_filter.pack(pady=20)

style = ttk.Style()
style.configure("Accent.TButton", font=("Arial", 12, "bold"))
