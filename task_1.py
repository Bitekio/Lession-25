"""Homework"""


import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path  
import pandas as pd


class ExcelReaderApp:
    """Класс для создания графического интерфейса приложения чтения файлов Excel."""

    def __init__(self, root):
        """
        Инициализация графического интерфейса.
        """
        self.root = root
        self.root.title("Excel Reader")

        self.path_label = tk.Label(root, text="Выберите папку с файлами Excel:")
        self.path_label.pack()

        self.path_entry = tk.Entry(root, width=50)
        self.path_entry.pack()

        self.browse_button = tk.Button(root, text="Выбрать папку", command=self.browse_folder)
        self.browse_button.pack()

        self.read_button = tk.Button(root, text="Прочитать файлы", command=self.read_files)
        self.read_button.pack()

        self.total_rows_label = tk.Label(root, text="")
        self.total_rows_label.pack()

    def browse_folder(self):
        """
        Функция для выбора папки с файлами Excel.
        """
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(tk.END, folder_path)

    def read_files(self):
        """
        Функция для чтения файлов Excel из выбранной папки и подсчета общего числа строк.
        """
        folder_path = self.path_entry.get()
        if not folder_path:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите папку.")
            return

        folder = Path(folder_path)  

        excel_files = [f for f in folder.iterdir() if f.suffix == '.xlsx']
        if not excel_files:
            messagebox.showerror("Ошибка", "Выбранная папка не содержит файлов Excel.")
            return

        total_rows = 0
        for file in excel_files:
            df = pd.read_excel(file)
            total_rows += len(df)

        self.total_rows_label.config(text=f"Общее количество строк: {total_rows}")

def main():
    """
    Основная функция для создания графического интерфейса и управления им.
    """
    root = tk.Tk()
    app = ExcelReaderApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
