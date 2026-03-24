import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import threading
import os
from pathlib import Path

class ADExportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Экспорт пользователей из AD группы")
        self.root.geometry("750x600")
        self.root.resizable(False, False)

        # Настройка стиля
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('Accent.TButton', font=('Arial', 10, 'bold'),
                             foreground='white', background='#0078d7',
                             bordercolor='#0078d7', focuscolor='none')
        self.style.map('Accent.TButton',
                       background=[('active', '#005a9e')])

        # Переменные
        self.group_var = tk.StringVar()
        self.file_var = tk.StringVar(value="ad_users.csv")
        self.recursive_var = tk.BooleanVar(value=True)
        self.append_var = tk.BooleanVar(value=False)
        self.encoding_var = tk.StringVar(value="UTF8")

        self.create_widgets()
        self.root.bind('<Return>', lambda event: self.export())

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовок
        title_label = ttk.Label(main_frame, text="Экспорт участников группы Active Directory",
                                 font=('Arial', 16, 'bold'))
        title_label.pack(pady=(0, 20))

        # Фрейм для полей ввода
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=10)

        # Поле "Имя группы"
        ttk.Label(input_frame, text="Имя группы AD:", font=('Arial', 10)).grid(
            row=0, column=0, sticky='w', pady=(0, 2))
        group_entry = ttk.Entry(input_frame, textvariable=self.group_var,
                                width=40, font=('Arial', 11))
        group_entry.grid(row=1, column=0, sticky='ew', padx=(0, 10), pady=(0, 10))
        group_entry.focus()

        # Поле "Путь к файлу"
        ttk.Label(input_frame, text="Выходной CSV-файл:", font=('Arial', 10)).grid(
            row=2, column=0, sticky='w', pady=(0, 2))
        file_frame = ttk.Frame(input_frame)
        file_frame.grid(row=3, column=0, sticky='ew', pady=(0, 10))
        file_entry = ttk.Entry(file_frame, textvariable=self.file_var,
                               font=('Arial', 11))
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        browse_btn = ttk.Button(file_frame, text="Обзор...", command=self.browse_file)
        browse_btn.pack(side=tk.RIGHT)

        # Опции (чекбоксы и выбор кодировки)
        options_frame = ttk.LabelFrame(input_frame, text="Параметры экспорта", padding="10")
        options_frame.grid(row=4, column=0, sticky='ew', pady=(10, 0))

        ttk.Checkbutton(options_frame, text="Рекурсивно (включая вложенные группы)",
                        variable=self.recursive_var).grid(row=0, column=0, sticky='w', padx=5)
        ttk.Checkbutton(options_frame, text="Добавлять в конец файла (append)",
                        variable=self.append_var).grid(row=1, column=0, sticky='w', padx=5)

        ttk.Label(options_frame, text="Кодировка CSV:").grid(row=0, column=1, sticky='w', padx=(20, 5))
        encoding_combo = ttk.Combobox(options_frame, textvariable=self.encoding_var,
                                       values=["UTF8", "UTF8BOM", "ASCII", "Unicode"],
                                       state="readonly", width=10)
        encoding_combo.grid(row=0, column=2, sticky='w')

        # Кнопки
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)

        export_btn = ttk.Button(button_frame, text="Экспортировать",
                                command=self.export, style='Accent.TButton')
        export_btn.grid(row=0, column=0, padx=10)

        reset_btn = ttk.Button(button_frame, text="Сбросить",
                               command=self.reset_fields)
        reset_btn.grid(row=0, column=1, padx=10)

        # Текстовое поле для результата
        result_frame = ttk.Frame(main_frame)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.result_text = tk.Text(
            result_frame, height=12, width=80, wrap='word',
            font=('Courier New', 9), relief='solid', borderwidth=1,
            highlightthickness=1, highlightcolor='#ccc', bg='white'
        )
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Запрещаем редактирование, но оставляем копирование
        self.result_text.bind('<Key>', self.block_key)
        self.result_text.bind('<Button-3>', self.show_context_menu)

        # Контекстное меню
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Копировать", command=self.copy_text)

        # Настройка сетки для input_frame
        input_frame.columnconfigure(0, weight=1)

    def block_key(self, event):
        if event.state & 0x4:  # Ctrl
            if event.keysym in ('c', 'C', 'a', 'A'):
                return
        if event.keysym in ('Left', 'Right', 'Up', 'Down', 'Home', 'End', 'Prior', 'Next'):
            return
        return 'break'

    def show_context_menu(self, event):
        self.context_menu.post(event.x_root, event.y_root)

    def copy_text(self):
        try:
            selected = self.result_text.get(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            selected = self.result_text.get(1.0, tk.END).strip()
        if selected:
            self.root.clipboard_clear()
            self.root.clipboard_append(selected)
            self.root.update()

    def browse_file(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=self.file_var.get()
        )
        if filename:
            self.file_var.set(filename)

    def reset_fields(self):
        self.group_var.set("")
        self.file_var.set("ad_users.csv")
        self.recursive_var.set(True)
        self.append_var.set(False)
        self.encoding_var.set("UTF8")
        self.result_text.delete(1.0, tk.END)
        self.root.focus_set()

    def check_ad_module(self):
        """Проверяет наличие модуля ActiveDirectory в PowerShell."""
        try:
            result = subprocess.run(
                ["powershell.exe", "-Command", "Get-Module -ListAvailable ActiveDirectory"],
                capture_output=True, text=True, timeout=10
            )
            return "ActiveDirectory" in result.stdout
        except Exception:
            return False

    def run_powershell(self, command):
        """Выполняет PowerShell-команду и возвращает (успех, сообщение)."""
        try:
            result = subprocess.run(
                ["powershell.exe", "-Command", command],
                capture_output=True, text=True, timeout=300  # 5 минут максимум
            )
            if result.returncode != 0:
                return False, result.stderr.strip()
            return True, result.stdout.strip()
        except subprocess.TimeoutExpired:
            return False, "Превышено время ожидания (5 минут)."
        except Exception as e:
            return False, str(e)

    def export(self):
        group = self.group_var.get().strip()
        file_path = self.file_var.get().strip()
        if not group or not file_path:
            messagebox.showerror("Ошибка", "Заполните имя группы и путь к файлу.")
            return

        # Проверка модуля AD
        if not self.check_ad_module():
            messagebox.showerror(
                "Модуль AD не найден",
                "Модуль ActiveDirectory не установлен в PowerShell.\n"
                "Убедитесь, что установлены средства управления AD."
            )
            return

        # Блокируем кнопку на время выполнения
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(1.0, "Выполняется экспорт...\n")
        self.result_text.update()

        # Запускаем в отдельном потоке
        threading.Thread(target=self._export_thread, args=(group, file_path), daemon=True).start()

    def _export_thread(self, group, file_path):
        """Фоновый поток для экспорта."""
        recursive = self.recursive_var.get()
        append = self.append_var.get()
        encoding = self.encoding_var.get()

        # Экранируем имя группы
        group_escaped = group.replace("'", "''")
        recursive_flag = "-Recursive" if recursive else ""
        append_flag = f"-Append:${str(append).lower()}"

        # Формируем PowerShell команду с добавлением поля Enabled
        ps_command = (
            f"Get-ADGroupMember -Identity '{group_escaped}' {recursive_flag} | "
            "ForEach-Object { "
            "    Get-ADUser -Filter {SamAccountName -eq $_.SamAccountName} -Properties Name, Mail, Enabled "
            "} | "
            "Select-Object SamAccountName, Name, Mail, Enabled | "
            f"Export-Csv -Path '{file_path}' -Encoding {encoding} -NoTypeInformation "
            f"{append_flag}"
        )

        # Выполняем
        success, error = self.run_powershell(ps_command)

        # Обновляем интерфейс в главном потоке, передаём кодировку
        self.root.after(0, self._export_callback, success, error, file_path, append, encoding)

    def _export_callback(self, success, error, file_path, append, encoding):
        """Обработка результата экспорта в главном потоке."""
        if not success:
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(1.0, f"Ошибка при экспорте:\n{error}")
            messagebox.showerror("Ошибка экспорта", f"Не удалось выполнить команду.\n\n{error}")
            return

        # Проверяем, создался ли файл
        path = Path(file_path)
        if not path.exists():
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(1.0, "Экспорт завершён, но файл не найден.")
            return

        # Определяем кодировку для чтения файла
        read_encoding = self._get_read_encoding(encoding)

        # Читаем первые строки для предварительного просмотра
        preview_lines = []
        try:
            with open(file_path, 'r', encoding=read_encoding) as f:
                # Пропускаем пустые строки в начале
                lines = [line for line in f if line.strip()]
                total_lines = len(lines)
                if total_lines > 0:
                    # Показываем заголовки и первые 5-10 строк (но не более 15)
                    preview_count = min(total_lines, 11)  # заголовок + 10 данных
                    preview_lines = lines[:preview_count]
        except Exception as e:
            preview_lines = [f"Ошибка чтения файла: {e}"]

        # Формируем отчёт
        result_text = self._format_report(
            group=self.group_var.get(),
            file_path=file_path,
            recursive=self.recursive_var.get(),
            append=self.append_var.get(),
            encoding=encoding,
            total_records=total_lines - 1 if total_lines > 0 else 0,
            preview=preview_lines
        )

        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(1.0, result_text)

    def _get_read_encoding(self, encoding):
        """Возвращает кодировку для чтения CSV в зависимости от выбранной."""
        if encoding == 'UTF8BOM':
            return 'utf-8-sig'
        elif encoding == 'ASCII':
            return 'ascii'
        elif encoding == 'Unicode':
            return 'utf-16'
        else:
            return 'utf-8'

    def _format_report(self, group, file_path, recursive, append, encoding, total_records, preview):
        """Формирует текст для отображения в текстовом поле."""
        lines = []
        lines.append("=" * 60)
        lines.append("Результат экспорта")
        lines.append("=" * 60)
        lines.append(f"Группа: {group}")
        lines.append(f"Файл: {file_path}")
        lines.append(f"Рекурсивно: {'Да' if recursive else 'Нет'}")
        lines.append(f"Добавление в файл: {'Да' if append else 'Нет'}")
        lines.append(f"Кодировка: {encoding}")
        lines.append(f"Всего записей (пользователей): {total_records}")
        lines.append("")

        if preview:
            lines.append("Предварительный просмотр (первые строки файла):")
            lines.append("-" * 60)
            # Добавляем номера строк
            for idx, line in enumerate(preview, 1):
                # Обрезаем длинные строки для удобства
                display_line = line.strip()
                if len(display_line) > 100:
                    display_line = display_line[:100] + "..."
                lines.append(f"{idx:3d}: {display_line}")
        else:
            lines.append("Нет данных для предварительного просмотра.")

        lines.append("")
        lines.append("Примечание: Файл содержит столбцы:")
        lines.append("  • Логин (SamAccountName)")
        lines.append("  • ФИО (Name)")
        lines.append("  • Email (Mail)")
        lines.append("  • Состояние (Enabled) – True = включена, False = отключена")

        return "\n".join(lines)

def main():
    root = tk.Tk()
    app = ADExportApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()