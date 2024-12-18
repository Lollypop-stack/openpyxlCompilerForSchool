import logging
import os
import tkinter as tk
from threading import Thread
from tkinter import messagebox, ttk, filedialog

from kundolukParser.grade import GradeProccesor
from kundolukParser.kparser import KParser


class Style:
    _mainFont = ("Arial", 14)
    _inputFont = ("Arial", 12)
    _labelFont = ("Arial", 14, "bold")

    _mainColor = "#3C3E52"
    _secondColor = "#F6F6F6"

    def __init__(
        self,
        mainFont=_mainFont,
        labelFont=_labelFont,
        inputFont=_inputFont,
        mainColor=_mainColor,
        secondColor=_secondColor,
    ) -> None:
        self.mainFont = mainFont
        self.labelFont = labelFont
        self.inputFont = inputFont
        self.mainColor = mainColor
        self.secondColor = secondColor


class GUI:
    def __init__(self, session: str, title="Kundoluk Parser") -> None:
        """
        Графический интерфейс для парсера

        Args:
            session (str): Сессия Кундолюка
            title (str, optional): Заголовок окна. Defaults to "Kundoluk Parser".
        """
        self.session = session

        self.logger = logging.getLogger(self.__class__.__name__)
        self.root = tk.Tk()
        self.root.title(title)
        self.style = Style()
        self.grade_entry = None
        self.quarter_entry = None
        self.directory = os.getcwd()
        self.directoryText = None

    def start(self) -> None:
        """
        Запуск графического интерфейса
        """
        self.logger.info("Запуск графического интерфейса")

        self._set_style()

        # Установка фоновых цветов для окна
        self.root.config(bg=self.style.secondColor)

        self._set_input()

        self._set_fileInput()

        # Кнопка с округленными углами и изменением стилей при наведении
        submit_button = ttk.Button(
            self.root,
            text="Отправить",
            style="RoundedButton.TButton",
            command=self._submit_with_preloader,
        )
        submit_button.pack(side="bottom", padx=0, pady=10)

        self.root.bind("<Return>", lambda ev: submit_button.invoke())

        self.root.mainloop()

    def _set_style(self) -> None:
        style = ttk.Style()
        style.configure(
            "RoundedButton.TButton",
            background=self.style.mainColor,  # Цвет фона по умолчанию
            foreground="#FFFFFF",  # Цвет текста по умолчанию
            font=self.style.mainFont,  # Кастомный шрифт
            padding=10,
            bd=2,
        )

        # Настройки для изменения кнопки при наведении
        style.map(
            "RoundedButton.TButton",
            background=[
                ("active", "#FFFFFF"),  # Цвет фона при наведении (светлый)
                ("!active", "#333333"),
            ],  # Цвет фона в неактивном состоянии (темный)
            foreground=[
                ("active", "#FFFFFF"),  # Цвет текста при наведении (темный)
                ("!active", "#333333"),
            ],
        ),  # Цвет текста в неактивном состоянии (белый)

        # Установка курсора для кнопки
        style.map(
            "RoundedButton.TButton", cursor=[("active", "hand2")]
        )  # Изменяет курсор на "hand2" (рука при наведении)

    def _set_input(self) -> None:
        frame = tk.Frame(self.root)
        frame.pack(side="top")
        # Заголовки
        tk.Label(
            frame,
            text="Введите класс (например, 4Б):",
            font=self.style.labelFont,
            bg=self.style.secondColor,
            fg=self.style.mainColor,
        ).grid(row=0, column=0, padx=10, pady=10)
        tk.Label(
            frame,
            text="Введите четверть (число):",
            font=self.style.labelFont,
            bg=self.style.secondColor,
            fg=self.style.mainColor,
        ).grid(row=1, column=0, padx=10, pady=10)

        # Поля ввода
        self.grade_entry = tk.Entry(
            frame, font=self.style.inputFont, bd=2, relief="solid", width=20
        )
        self.quarter_entry = tk.Entry(
            frame, font=self.style.inputFont, bd=2, relief="solid", width=20
        )

        self.grade_entry.grid(row=0, column=1, padx=10, pady=10)
        self.quarter_entry.grid(row=1, column=1, padx=10, pady=10)

    def _set_fileInput(self) -> None:
        frame = tk.Frame(self.root)
        frame.pack(side="top", fill="x")

        import_button = tk.Button(
            frame,
            text="Выберите директорию",
            command=self._import_dir,
        )
        import_button.pack(side="left", padx=30, pady=5)

        self.directoryText = tk.Label(
            frame,
            text=self.directory,
            bg=self.style.secondColor,
            fg=self.style.mainColor,
            justify="left",
        )
        self.directoryText.pack(side="left", pady=5)

    def _import_dir(self) -> None:
        dir_path = filedialog.askdirectory(title="Выберите директорию")
        if dir_path:
            self.logger.info(f"Выбрана директория: {dir_path}")
            self.directory = dir_path
            self.directoryText.config(text=dir_path)

    def center_window(self, window, width, height) -> None:
        """Центрирование окна на экране"""
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()

        position_top = int(screen_height / 2 - height / 2)
        position_right = int(screen_width / 2 - width / 2)

        window.geometry(f"{width}x{height}+{position_right}+{position_top}")

    def show_preloader(self) -> tuple:
        # Здесь создаем окно прелоадера с зеленой полосой загрузки
        preloader = tk.Toplevel()
        preloader.title("Загрузка...")
        self.center_window(preloader, 300, 120)  # Ширина 300, высота

        # Элемент текста
        label = tk.Label(
            preloader,
            text="Пожалуйста, подождите...",
            font=self.style.inputFont,
            bg=self.style.secondColor,
            fg=self.style.mainColor,
        )
        label.pack(padx=20, pady=10)

        # Зеленая полоска загрузки
        progress = ttk.Progressbar(
            preloader,
            length=250,
            mode="indeterminate",
            style="TProgressbar",
        )
        progress.pack(padx=20, pady=10)

        # Начинаем анимацию полосы
        progress.start()

        return preloader, progress

    def _perform_task(self, grade, quarter, preloader):
        try:
            parser = KParser(self.session)
            grades = dict(parser.gradesList)
            gradeId = grades.get(grade)
            if gradeId and quarter > 0:
                grade = parser.get_grade((grade, gradeId), quarter)
                xlsxFile = grade.to_excel(self.directory)
                GradeProccesor(xlsxFile).start()
            else:
                print("Неверный ввод. Попробуйте снова.")
        except Exception as e:
            self.logger.warning(f"Ошибка: {e}. Попробуйте снова.", exc_info=e)
        finally:
            preloader.destroy()  # Закрываем прелоадер после завершения

    def _submit_with_preloader(self):
        """
        Обработчик кнопки "Отправить" с прелоадером.
        """
        grade = self.grade_entry.get().strip().upper()
        quarter = self.quarter_entry.get().strip()

        if not grade or not quarter.isdigit():
            messagebox.showerror("Ошибка", "Введите корректные значения!")
            return

        # Создаем окно прелоадера
        preloader, progress = self.show_preloader()

        # Выполняем задачу в отдельном потоке
        task_thread = Thread(
            target=self._perform_task, args=(grade, int(quarter), preloader)
        )
        task_thread.start()
