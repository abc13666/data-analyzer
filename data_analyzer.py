import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np

class DataAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализатор данных")
        self.root.geometry("700x550")
        
        self.df = None
        self.current_file = None
        
        # --- Инструменты ---
        self.tab_control = ttk.Notebook(root)
        
        self.tab_load = ttk.Frame(self.tab_control)
        self.tab_analyze = ttk.Frame(self.tab_control)
        self.tab_filter = ttk.Frame(self.tab_control)
        
        self.tab_control.add(self.tab_load, text='📁 Импорт')
        self.tab_control.add(self.tab_analyze, text='📊 Анализ')
        self.tab_control.add(self.tab_filter, text='🔍 Фильтр')
        self.tab_control.pack(expand=1, fill='both')
        
        self._setup_load_tab()
        self._setup_analyze_tab()
        self._setup_filter_tab()
        
    def _setup_load_tab(self):
        """Настройка вкладки импорта"""
        frame = ttk.LabelFrame(self.tab_load, text="Добавить файл")
        frame.pack(pady=20, padx=20, fill='x')
        
        ttk.Label(frame, text="Состояние файла:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        self.status_label = ttk.Label(frame, text="Опустошено", font=("Arial", 12, "bold"))
        self.status_label.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Button(frame, text="Загрузить CSV/Excel", command=self.load_file).grid(row=1, column=0, columnspan=2, padx=10, pady=10)
        
        # Таблица для показов данных
        self.tree = ttk.Treeview(self.tab_load, columns=('ID', 'Col1', 'Col2', 'Col3'), show='headings')
        self.tree.heading('ID', text='ID')
        self.tree.heading('Col1', text='Колонка 1')
        self.tree.heading('Col2', text='Колонка 2')
        self.tree.heading('Col3', text='Колонка 3')
        
        self.tree.column('ID', width=80)
        self.tree.column('Col1', width=150)
        self.tree.column('Col2', width=150)
        self.tree.column('Col3', width=150)
        
        self.tree.pack(fill='both', expand=True, padx=20, pady=10)
        
    def _setup_analyze_tab(self):
        """Настройка вкладки анализа"""
        frame = ttk.LabelFrame(self.tab_analyze, text="Статистика")
        frame.pack(pady=20, padx=20, fill='both', expand=True)
        
        ttk.Label(frame, text="Статистика по колонкам (Среднее, Медиана, Макс, Мин)").pack(pady=5)
        
        # Играемся с акциями
        self.analyze_tree = ttk.Treeview(frame, columns=('Column', 'Mean', 'Median', 'Max', 'Min'), show='headings')
        self.analyze_tree.heading('Column', text='Колонка')
        self.analyze_tree.heading('Mean', text='Среднее')
        self.analyze_tree.heading('Median', text='Медиана')
        self.analyze_tree.heading('Max', text='Макс')
        self.analyze_tree.heading('Min', text='Мин')
        
        self.analyze_tree.column('Column', width=100)
        self.analyze_tree.column('Mean', width=100)
        self.analyze_tree.column('Median', width=100)
        self.analyze_tree.column('Max', width=100)
        self.analyze_tree.column('Min', width=100)
        
        self.analyze_tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        ttk.Button(frame, text="Показать статистику", command=self.show_stats).pack(pady=5)
        
    def _setup_filter_tab(self):
        """Настройка вкладки фильтра"""
        # Основной фрейм
        frame = ttk.LabelFrame(self.tab_filter, text="Фильтрация")
        frame.pack(pady=20, padx=20, fill='both', expand=True)
        
        # --- Верхняя часть: поля ввода (используем grid) ---
        input_frame = ttk.Frame(frame)
        input_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(input_frame, text="Сумма >=").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.threshold_entry = ttk.Entry(input_frame, width=15)
        self.threshold_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(input_frame, text="Результаты фильтра", command=self.filter_data).grid(row=0, column=2, padx=5, pady=5)
        
        # --- Нижняя часть: таблица и кнопка (используем pack) ---
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.filter_tree = ttk.Treeview(table_frame, columns=('ID', 'Col1', 'Col2', 'Col3'), show='headings')
        self.filter_tree.heading('ID', text='ID')
        self.filter_tree.heading('Col1', text='Колонка 1')
        self.filter_tree.heading('Col2', text='Колонка 2')
        self.filter_tree.heading('Col3', text='Колонка 3')
        
        self.filter_tree.column('ID', width=80)
        self.filter_tree.column('Col1', width=150)
        self.filter_tree.column('Col2', width=150)
        self.filter_tree.column('Col3', width=150)
        
        self.filter_tree.pack(fill='both', expand=True)
        
        ttk.Button(table_frame, text="Export в Excel", command=self.export_filtered).pack(pady=5)
        
    def load_file(self):
        """Загрузка файла"""
        file_path = filedialog.askopenfilename(
            title="Выберите файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    # Гибкое чтение CSV: если в файле разделитель ';', pandas сам его найдет
                    self.df = pd.read_csv(file_path, sep=None, engine='python', encoding='utf-8')
                else:
                    self.df = pd.read_excel(file_path)
                
                self.current_file = file_path
                self.status_label.config(text="✅ Файл загружен")
                
                # Очищаем таблицу ПРАВИЛЬНО
                for item in self.tree.get_children():
                    self.tree.delete(item)
                
                # Показываем первые 5 строк
                rows = self.df.head(5).to_numpy()
                for i, row in enumerate(rows):
                    self.tree.insert('', 'end', values=(i+1, *row))
                
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))
        
    def show_stats(self):
        """Показать статистику"""
        if self.df is None or self.df.empty:
            messagebox.showerror("Ошибка", "Нет данных для анализа!")
            return
        
        # Очищаем таблицу
        for item in self.analyze_tree.get_children():
            self.analyze_tree.delete(item)
        
        # Показываем статистику
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        for col in numeric_cols:
            mean = self.df[col].mean()
            median = self.df[col].median()
            max_val = self.df[col].max()
            min_val = self.df[col].min()
            
            self.analyze_tree.insert('', 'end', values=(col, mean, median, max_val, min_val))
        
    def filter_data(self):
        """Фильтрация данных"""
        if self.df is None or self.df.empty:
            messagebox.showerror("Ошибка", "Нет данных!")
            return
        
        threshold = self.threshold_entry.get().strip()
        if not threshold:
            messagebox.showerror("Ошибка", "Введите порог!")
            return
        
        try:
            threshold = float(threshold)
        except:
            messagebox.showerror("Ошибка", "Порог должен быть числом!")
            return
        
        # Если в данных есть числовая колонка, фильтруем по ней
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) == 0:
            messagebox.showerror("Ошибка", "Нет числовых колонок!")
            return
        
        # Фильтрация по первой числовой колонке
        col = numeric_cols[0]
        filtered = self.df[self.df[col] >= threshold]
        
        # Показываем результат
        for item in self.filter_tree.get_children():
            self.filter_tree.delete(item)
        
        rows = filtered.head(10).to_numpy()
        for i, row in enumerate(rows):
            self.filter_tree.insert('', 'end', values=(i+1, *row))
        
    def export_filtered(self):
        """Экспорт данных в Excel"""
        if self.df is None or self.df.empty:
            messagebox.showerror("Ошибка", "Нет данных!")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Экспорт в Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if file_path:
            try:
                self.df.to_excel(file_path, index=False)
                messagebox.showinfo("Успех", "Файл сохранен!")
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = DataAnalyzer(root)
    root.mainloop()