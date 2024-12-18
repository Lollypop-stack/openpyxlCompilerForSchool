import logging
import time
from io import StringIO
from threading import Thread

import pandas as pd
import requests
from bs4 import BeautifulSoup
from pandas import DataFrame

from kundolukParser.grade import GradeProccesor, Grade, Subject


class RThread(Thread):
    """
    Расширение потока для выполнения задачи и хранения результата.
    """

    def __init__(self, target, args):
        Thread.__init__(self)
        self.target = target
        self.result = None

    def run(self) -> None:
        self.result = self.target()


class KParser:
    # Список классов и их идентификаторов
    gradesList = [
        ("2А", 90359),
        ("2Б", 90360),
        ("3А", 90340),
        ("3Б", 90341),
        ("4А", 90343),
        ("4Б", 90344),
        ("5А", 90345),
        ("5Б", 90346),
        ("6А", 90347),
        ("6Б", 90348),
        ("7А", 90349),
        ("7Б", 90350),
        ("8А", 90351),
        ("8Б", 90352),
        ("9А", 90353),
        ("9Б", 90354),
        ("10А", 90355),
        ("10Б", 90356),
        ("11А", 90357),
        ("11Б", 90358),
        ("2В", 90361),
    ]

    def __init__(self, session: str) -> None:
        """
        Парсер Кундолюка

        Args:
            session (str): Сессия Кундолюка
        """
        self.logger = logging.getLogger(self.__class__.__name__)
        self.__headers = {
            "cookie": f"language=ru; session={session}; deferApp=true; string=670fbb44e3637",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.6324.206 Safari/537.36",
        }
        self.url = "https://kundoluk.edu.kg/journal2"
        self.s = requests.Session()

    def get_grade(self, grade: tuple[str, int], quarter: int = 1) -> Grade:
        """
        Получение данных для конкретного класса и четверти.

        Args:
            grade (tuple[str, int]): Класс и его ID
            quarter (int): Номер четверти

        Returns:
            Grade: Объект класса
        """
        try:
            self.logger.info(f"Начало получения {grade[0]} класса")

            querystring = {"class": str(grade[1]), "quarter": str(quarter)}
            response = self.s.get(self.url, headers=self.__headers, params=querystring)
            response.raise_for_status()  # Проверка успешности запроса

            bs = BeautifulSoup(response.content, "lxml")
            subjects = bs.find("ul", class_="uk-subnav").find_all("a")
            threads = []
            for link in subjects:
                url = link["href"]
                name = link.text.strip()

                self.logger.info(f"Парсинг предмета: {name}")
                # Создаем поток для получения данных о предмете
                thread = RThread(target=lambda: self.get_subject(url, name), args=())
                thread.start()
                threads.append(thread)

                time.sleep(0.1)  # Небольшая пауза между запросами

            subjects = []
            for thread in threads:
                while thread.result is None:  # Ожидание завершения потока
                    time.sleep(0.1)
                subject = thread.result
                if type(subject) != int:
                    subjects.append(subject)

            # Сортировка по названию урока
            subjects.sort(key=lambda x: x.name)

            self.logger.info(f"Конец получения {grade[0]} класса")

            return Grade(subjects, grade[0], quarter)
        except requests.exceptions.RequestException as e:
            raise Exception(f"Ошибка при запросе данных для класса {grade[0]}: {e}")

    def get_subject(self, url: str, name: str) -> Subject:
        """
        Получение данных о предмете по ссылке.

        Args:
            url (str): Ссылка на предмет
            name (str): Название предмета

        Returns:
            Subject: Объект урока
        """
        response = self.s.request("GET", url, headers=self.__headers)
        # Проверка успешности запроса
        if not response.ok:
            raise Exception(
                "subjectError:",
                f"subject url({url}) response code {response.status_code}",
            )

        bs = BeautifulSoup(response.content, "lxml")

        try:
            htmlTable = bs.find(
                "table", class_="elementFixed-striped"
            )  # Таблица с данными
            # Удаляем лишние элементы из таблицы
            trash = htmlTable.find_all("span", class_="uk-margin-xsmall-right")
            [i.extract() for i in trash]

            # Конвертируем в DataFrame
            table = pd.read_html(StringIO(str(htmlTable)))[0]
        except Exception:
            self.logger.warning(f"Отсутствуют данные предмета: {name}")
            return 10
        return Subject(name, table)

    @classmethod
    def magic(cls) -> Grade:
        """
        Обработка пользовательского ввода в консоли

        Returns:
            Grade: Объект класса
        """
        try:
            while True:
                try:
                    inpGrade = input("Введите класс(Пример: 4Б): ").upper()
                    inpQuarter = int(input("Введите четверть: "))
                    grades = dict(cls.gradesList)
                    gradeId = grades.get(inpGrade)
                    if gradeId and inpQuarter > 0:
                        grade = cls.get_grade((inpGrade, gradeId), inpQuarter)
                        xlsxFile = grade.to_excel()
                        GradeProccesor(xlsxFile).start()

                        return grade
                    else:
                        print("Неверный ввод. Попробуйте снова.")
                except Exception as e:
                    cls.logger.warning(f"Ошибка: {e}. Попробуйте снова.", exc_info=e)
        except KeyboardInterrupt:
            cls.logger.warning("Program terminated by user")
