import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys


class ExcelToCSVConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Конвертер Excel в CSV")
        self.root.geometry("800x600")

        style = ttk.Style()
        style.configure("TButton", padding=6)
        style.configure("TLabel", padding=6)
        style.configure("TCombobox", padding=6)

        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.label = ttk.Label(
            self.main_frame, text="Выберите файл Excel", font=("Arial", 14))
        self.label.pack(expand=True)

        self.button_select = ttk.Button(
            self.main_frame, text="Выбрать файл", command=self.select_file)
        self.button_select.pack(pady=5)

        self.preview_button = ttk.Button(
            self.main_frame, text="Предварительный просмотр", command=self.preview_data)
        self.preview_button.pack(pady=5)

        self.preview_frame = ttk.Frame(self.main_frame)
        self.preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(self.preview_frame, show='headings')
        self.tree.pack(expand=True, fill='both')

    def select_file(self):
        file_path = filedialog.askopenfilename(title="Выберите файл Excel", filetypes=[
            ("Excel файлы", "*.xls;*.xlsx;*.xlsm;*.xlsb")])
        if file_path:
            output_dir = filedialog.askdirectory(
                title="Выберите директорию для сохранения CSV")
            if output_dir:
                try:
                    self.convert_excel_to_csv(file_path, output_dir)
                except Exception as e:
                    messagebox.showerror(
                        "Ошибка", f"Ошибка при загрузке файла: {e}")

    def convert_excel_to_csv(self, input_file, output_dir):
        if not os.path.isfile(input_file):
            raise FileNotFoundError(f"Файл не найден: {input_file}")

        base_name = os.path.splitext(os.path.basename(input_file))[0]
        output_file = os.path.join(output_dir, f"{base_name}.csv")

        try:
            df = self.load_excel(input_file)
            if df.empty:
                raise ValueError("Файл пуст или не содержит данных.")

            columns_to_drop = self.check_columns(df)
            df.drop(columns=columns_to_drop, inplace=True)

            filtered_df = self.filter_dataframe(df)

            self.insert_new_columns(filtered_df)


            if filtered_df.shape[1] >= 6:
                filtered_df.iloc[:, 4] = pd.to_datetime(filtered_df.iloc[:, 4], errors='coerce').dt.date
                filtered_df.iloc[:, 5] = pd.to_datetime(filtered_df.iloc[:, 5], errors='coerce').dt.date

            # Удаление лишних столбцов, если они есть
            if filtered_df.shape[1] > 8:
                filtered_df.drop(columns=filtered_df.columns[8], inplace=True, errors='ignore')
            if filtered_df.shape[1] > 9:
                filtered_df.drop(columns=filtered_df.columns[8], inplace=True, errors='ignore')

            filtered_df.to_csv(output_file, index=False,
                               header=False, encoding='windows-1251', sep=';')
            messagebox.showinfo(
                "Успех", f"Файл успешно конвертирован в: {output_file}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при конвертации: {e}")

    def load_excel(self, input_file):
        if input_file.endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')):
            return pd.read_excel(input_file, engine='openpyxl')
        elif input_file.endswith('.xlsb'):
            return pd.read_excel(input_file, engine='pyxlsb')
        else:
            raise ValueError("Неподдерживаемый формат файла.")


    def check_columns(self, df):
        values_to_check = ['Сумма НДС (руб)', 'Сумма с НДС (руб.)']
        columns_to_drop = []
        for value in values_to_check:
            if (df.iloc[:10].astype(str).apply(lambda x: x.str.contains(value)).any(axis=1)).any():
                columns_to_drop.extend(df.columns[df.isin([value]).any()])
        return list(set(columns_to_drop))

    def filter_dataframe(self, df):
        return df[df.iloc[:, 0].apply(lambda x: isinstance(x, (int, float, str)) and str(x).isdigit())]

    def insert_new_columns(self, filtered_df):
        project_type_map = {
            1: 'Проектная документация',
            2: 'Рабочая документация',
            3: 'Обмерные обследовательские работы',
            4: 'Инженерные изыскания',
            5: 'Основные проектные решения',
            6: 'Разработка документации по планировке территории',
            7: 'Исходно-разрешительная документация',
            8: 'Авторский надзор',
            9: 'Обоснование инвестиций, ТЭО',
            10: 'Прочие работы/услуги'
        }

        new_column_values = [
            project_type_map.get(int(row[0]), '') if str(row[0]).isdigit() else '' for _, row in filtered_df.iterrows()
        ]

        num_rows = len(filtered_df)
        if len(new_column_values) < num_rows:
            new_column_values += [''] * (num_rows - len(new_column_values))
        elif len(new_column_values) > num_rows:
            new_column_values = new_column_values[:num_rows]

        etap_column_index = self.get_etap_column_index(filtered_df)

        filtered_df.insert(etap_column_index + 1, 'Новый столбец', new_column_values)
        filtered_df.insert(1, 'Пустой столбец', [''] * len(filtered_df))
        filtered_df.insert(7, 'H', ['20%'] * len(filtered_df))

    def get_etap_column_index(self, filtered_df):
        etap_index = filtered_df.apply(lambda row: next((i for i, cell in enumerate(row) if str(cell).startswith('Этап')), None), axis=1).dropna().astype(int).min()
        return etap_index if pd.notna(etap_index) else 0

    def update_treeview(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)


        for col in df.columns:
            self.tree.heading(col, text="")
            self.tree.column(col, anchor="center", stretch=True, width=100)

        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def preview_data(self):
        file_path = filedialog.askopenfilename(title="Выберите файл Excel", filetypes=[
            ("Excel файлы", "*.xls;*.xlsx;*.xlsm;*.xlsb")])
        if file_path:
            try:
                df = self.load_excel(file_path)
                if df.empty:
                    raise ValueError("Файл пуст или не содержит данных.")

                columns_to_drop = self.check_columns(df)
                df.drop(columns=columns_to_drop, inplace=True)
                filtered_df = self.filter_dataframe(df)
                self.insert_new_columns(filtered_df)

                if filtered_df.shape[1] >= 6:
                    filtered_df.iloc[:, 4] = pd.to_datetime(filtered_df.iloc[:, 4], errors='coerce').dt.date
                    filtered_df.iloc[:, 5] = pd.to_datetime(filtered_df.iloc[:, 5], errors='coerce').dt.date

                if filtered_df.shape[1] > 8:
                    filtered_df.drop(columns=filtered_df.columns[8], inplace=True, errors='ignore')
                if filtered_df.shape[1] > 9:
                    filtered_df.drop(columns=filtered_df.columns[8], inplace=True, errors='ignore')

                self.update_treeview(filtered_df)
            except Exception as e:
                messagebox.showerror(
                    "Ошибка", f"Ошибка при загрузке файла: {e}")

def main():
    root = tk.Tk()
    app = ExcelToCSVConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
