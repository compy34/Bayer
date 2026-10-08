import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from convertor import build_report_dataframe, generate_excel_report
from Data.Controller import get_pharmacy_data, load_requisites_file


class ReportApp:

    def __init__(self, root):
        self.root = root
        self.root.title(
            'Генератор маркетингових звітів та актів (Автоматизація)'
        )
        self.root.geometry('1100x700')
        self.root.minsize(900, 600)
        self.df = None
        self.df_req = None
        self.requisites_file_var = tk.StringVar(value='')
        self.pharmacy_name_var = tk.StringVar(value='')
        self._build_ui()
        self._load_data_async()

    def _build_ui(self):
        root_frame = ttk.Frame(self.root, padding=12)
        root_frame.pack(fill='both', expand=True)
        root_frame.columnconfigure(1, weight=1)
        root_frame.rowconfigure(5, weight=1)

        ttk.Label(
            root_frame,
            text='Генератор звітів та актів',
            font=('TkDefaultFont', 15, 'bold'),
        ).grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 14))

        ttk.Label(root_frame, text='Файл реквізитів:').grid(
            row=1, column=0, sticky='w', pady=4
        )
        ttk.Entry(
            root_frame,
            textvariable=self.requisites_file_var,
            state='readonly',
            width=75,
        ).grid(row=1, column=1, sticky='ew', pady=4)
        ttk.Button(
            root_frame, text='Обрати файл...', command=self.select_requisites_file
        ).grid(row=1, column=2, sticky='w', padx=(8, 0), pady=4)

        ttk.Label(root_frame, text='Мережа / Клієнт:').grid(
            row=2, column=0, sticky='w', pady=4
        )
        self.pharmacy_combo = ttk.Combobox(
            root_frame,
            textvariable=self.pharmacy_name_var,
            state='readonly',
            width=75,
        )
        self.pharmacy_combo.grid(row=2, column=1, columnspan=2, sticky='ew', pady=4)

        ttk.Label(root_frame, text='Період (Звітний):').grid(
            row=3, column=0, sticky='w', pady=4
        )
        period_frame = ttk.Frame(root_frame)
        period_frame.grid(row=3, column=1, columnspan=2, sticky='w', pady=4)

        self.start_var = tk.StringVar(value='01.08.2026')
        ttk.Entry(period_frame, textvariable=self.start_var, width=18).pack(
            side='left'
        )
        ttk.Label(period_frame, text=' по ').pack(side='left')
        self.end_var = tk.StringVar(value='31.08.2026')
        ttk.Entry(period_frame, textvariable=self.end_var, width=18).pack(
            side='left'
        )

        buttons = ttk.Frame(root_frame)
        buttons.grid(row=4, column=0, columnspan=3, sticky='w', pady=(12, 8))
        ttk.Button(buttons, text='Попередній перегляд', command=self.preview).pack(
            side='left', padx=(0, 8)
        )
        ttk.Button(buttons, text='Завантажити Excel', command=self.download).pack(
            side='left'
        )

        self.status_var = tk.StringVar(value='Завантаження даних з БД...')
        ttk.Label(root_frame, textvariable=self.status_var).grid(
            row=5, column=0, columnspan=3, sticky='w', pady=(0, 8)
        )

        table_frame = ttk.Frame(root_frame)
        table_frame.grid(row=6, column=0, columnspan=3, sticky='nsew')
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_frame, show='headings')
        self.tree.grid(row=0, column=0, sticky='nsew')

        y_scroll = ttk.Scrollbar(
            table_frame, orient='vertical', command=self.tree.yview
        )
        y_scroll.grid(row=0, column=1, sticky='ns')
        x_scroll = ttk.Scrollbar(
            table_frame, orient='horizontal', command=self.tree.xview
        )
        x_scroll.grid(row=1, column=0, sticky='ew')
        self.tree.configure(
            yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set
        )

    def _load_data_async(self):
        def load():
            data = get_pharmacy_data()
            self.root.after(0, lambda: self._data_loaded(data))

        threading.Thread(target=load, daemon=True).start()

    def _data_loaded(self, data):
        self.df = data
        if data is None or data.empty:
            self.status_var.set('Помилка завантаження з БД або дані порожні.')
        else:
            self.status_var.set(
                f'Дані з БД успішно завантажено. Рядків: {len(data)}'
            )

    def _update_pharmacy_combo(self):
        """Заповнює комбобокс мережами з файлу реквізитів"""
        if self.df_req is not None and not self.df_req.empty:
            target_col = None
            for col_candidate in [
                'chain_name',
                'Клиент из базы данных',
                'Мережа',
                'Назва',
            ]:
                if col_candidate in self.df_req.columns:
                    target_col = col_candidate
                    break

            if target_col:
                chains = (
                    self.df_req[target_col].dropna().astype(str).str.strip()
                )
                unique_chains = sorted(
                    c for c in chains.unique() if c and c.lower() != 'nan'
                )
                self.pharmacy_combo['values'] = ('Усі аптеки', *unique_chains)
                self.pharmacy_combo.current(0)
                self.status_var.set(
                    f'Завантажено мереж з реквізитів: {len(unique_chains)}'
                )

    def _filtered_pharmacy_data(self):
        if self.df is None or self.df.empty:
            raise ValueError('Дані з бази відсутні.')
        selected_name = self.pharmacy_name_var.get().strip()
        if not selected_name or selected_name == 'Усі аптеки':
            return self.df

        # Фільтруємо дані за назвою мережі / аптеки
        filtered = self.df[
            self.df['PHARMACY_NAME'].fillna('').astype(str).str.strip().str.lower()
            == selected_name.lower()
        ]
        if filtered.empty:
            # Якщо точного збігу в базі немає, повертаємо весь датасет для подальшої обробки
            return self.df
        return filtered

    def select_requisites_file(self):
        filename = filedialog.askopenfilename(
            title='Виберіть файл реквізитів',
            initialdir=str(Path.cwd()),
            filetypes=[
                ('Усі таблиці', '*.csv *.xlsx *.xls'),
                ('CSV файли', '*.csv'),
                ('Excel файли', '*.xlsx *.xls'),
            ],
        )
        if not filename:
            return
        try:
            requisites = load_requisites_file(filename)
            if requisites.empty:
                raise ValueError('Файл реквізитів порожній.')
        except (OSError, ValueError, UnicodeError) as error:
            messagebox.showerror('Помилка', str(error))
            return

        self.df_req = requisites
        self.requisites_file_var.set(filename)
        self._update_pharmacy_combo()

    def _validate_period(self):
        try:
            start = datetime.strptime(self.start_var.get().strip(), '%d.%m.%Y')
            end = datetime.strptime(self.end_var.get().strip(), '%d.%m.%Y')
        except ValueError as error:
            raise ValueError('Невірний формат дати. Використовуйте ДД.ММ.РРРР.') from error
        if start > end:
            raise ValueError('Дата початку не може бути більшою за дату кінця.')
        return self.start_var.get().strip(), self.end_var.get().strip()

    def _selected_requisites(self):
        if self.df_req is None or self.df_req.empty:
            raise ValueError('Будь ласка, оберіть файл реквізитів!')
        return self.df_req

    def _get_preview_data(self):
        self._selected_requisites()
        self._validate_period()
        report = build_report_dataframe(self._filtered_pharmacy_data())
        if report is None or report.empty:
            report = build_report_dataframe(self.df)
        if report is None or report.empty:
            raise ValueError('Немає даних для формування звіту.')
        return report

    def preview(self):
        try:
            report = self._get_preview_data()
        except (ValueError, KeyError) as error:
            messagebox.showwarning('Попередження', str(error))
            return

        self.tree.delete(*self.tree.get_children())
        columns = [str(column) for column in report.columns]
        self.tree['columns'] = columns
        for column in columns:
            self.tree.heading(column, text=column)
            self.tree.column(column, width=150, minwidth=80, anchor='center')

        for row in report.head(200).itertuples(index=False, name=None):
            self.tree.insert(
                '', 'end', values=tuple(str(value) for value in row)
            )

        self.status_var.set(
            f'Відображено рядків: {min(len(report), 200)} з {len(report)}'
        )

    def download(self):
        try:
            self._get_preview_data()
        except (ValueError, KeyError) as error:
            messagebox.showwarning('Попередження', str(error))
            return

        filename = filedialog.asksaveasfilename(
            title='Зберегти звіт як',
            initialdir=str(Path.cwd()),
            initialfile='marketing_report.xlsx',
            defaultextension='.xlsx',
            filetypes=[('Excel файли', '*.xlsx')],
        )
        if not filename:
            return

        try:
            start, end = self._validate_period()
            generate_excel_report(
                self._filtered_pharmacy_data(),
                self._selected_requisites(),
                filename,
                period_start=start,
                period_end=end,
                selected_pharmacy=self.pharmacy_name_var.get(),
            )
        except (OSError, ValueError, KeyError) as error:
            messagebox.showerror('Помилка', str(error))
            return

        self.status_var.set(f'Звіт збережено: {filename}')
        messagebox.showinfo('Успіх', f'Файл успішно збережено:\n{filename}')


def main():
    root = tk.Tk()
    ReportApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()