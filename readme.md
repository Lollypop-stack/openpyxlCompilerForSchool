# Kundoluk Parser

Этот проект представляет собой инструмент для обработки и анализа данных в формате Excel с результатами учеников. Он использует библиотеку `tkinter` для создания графического интерфейса пользователя (GUI) и предоставляет функциональность для подсчета средних баллов, категоризации студентов и генерации диаграмм. Программа также поддерживает работу с многозадачностью, используя потоки для выполнения долгих операций без блокировки пользовательского интерфейса.

## Функции

- **Ввод данных**: Пользователь вводит класс (например, "4Б") и четверть (число) для обработки соответствующих данных.

- **Прелоадер**: Визуальный индикатор прогресса, показывающий пользователю, что данные обрабатываются.

- **Обработка данных**:
  - Программа анализирует данные из Excel-файлов.
  - Вычисляет средний балл для каждого студента.
  - Классифицирует студентов в категории (Отл., Уд., Тр., Дв., Нз.) в зависимости от их среднего балла.
  - Генерирует круговую диаграмму, отображающую распределение студентов по категориям.

- **Графический интерфейс**:
  - Простое окно для ввода данных.
  - Кнопка для отправки данных на обработку.
  - Стильный дизайн с кнопками и полями ввода, которые меняют свой внешний вид при наведении.


## Структура проекта
- **KParser** - Класс для работы с системой Kundoluk.
 ### Основные функции:
  - Получение данных о классах (get_grade)
  - Получение данных по предметам (get_subject)
  - Обработка пользовательского ввода (magic)

- **Grade** - Создание Excel файла с данными по предметам в указанном пути.
 ### Основные функции:
  - Создает **Excel** таблицу исходя из данных по предметам, взятых со страницы класса на сайте Kundoluk.

- **assign_categories** - Добавляет категории (Отл., Уд., Тр., Дв., Нз.) рядом со средним баллом.
 ### Аргументы:
  - **result_sheet** -- лист Excel с результатами.
  - **num_subjects** -- количество предметов.

- **calculate_averages** - Создает лист Result, в котором хранятся данные о средних баллах по каждому из предметов для каждого ученика. Так же высчитывается средний балл по всем предметам. В зависимости от итогового балла добавляется приписка "Отл."/"Уд."/"Тр."/"Дв.".
 ### Возможности:
  - Добавляет колонку с итоговым баллом за четверть, а также рядом добавляется уровень успеваемости.
  - Создает мини-таблицу с процентным соотношением успеваемости класса, а также добавляется визуальная составляющая(диаграмма).

- **start_ui** - Блок кода, создающий пользовательский интерфейс. Он минималистичен, прост и удобен в использовании.
 ### Что еще внутри интересного?:
  - Добавлен прелоадер, отображающий процесс создания файла
  - Все окна отцентрированны посередине экрана благодаря встроенной функции. Это сделано для удобства использования.

# Важно!
## **Перед запуском приложения войдите в систему Kundoluk на устройстве, на котором будете запускать программу!**
## А также убедитесь, что файлы Cookie разрешены в вашем браузере.


## Установка

1. Скачайте zip файл:

2. Установите необходимые библиотеки:

    ```bash
    pip install -r requirements.txt
    ```

## Запуск

Для запуска программы просто выполните следующий скрипт:

```bash
python main.py
