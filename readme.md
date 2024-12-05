# Kundoluk Parser

Этот проект представляет собой инструмент для извлечения оценок учеников из системы `Kundoluk`, обработки этих данных и сохранения их в Excel-формате с расчетом среднего балла для каждого ученика по предметам. Программа поддерживает как консольный интерфейс, так и графический интерфейс с использованием библиотеки `tkinter`.

## Описание

### Классы и функции

1. **Класс `RThread`**
   - Этот класс наследует от `Thread` и используется для выполнения многозадачных операций, таких как извлечение данных о каждом предмете в отдельном потоке.
   
2. **Класс `Grade`**
   - Этот класс хранит информацию о классе, четверти и предмете. Он позволяет экспортировать данные в Excel и выводить таблицу с оценками.

3. **Класс `KParser`**
   - Основной класс, который выполняет все необходимые операции: получение данных о классе и четверти, обработка оценок по каждому предмету и генерация результирующего Excel файла.

4. **Функция `calculate_averages`**
   - Обрабатывает данные и записывает средние баллы для каждого ученика по каждому предмету в новый Excel файл.

5. **Функция `extract_digit`**
   - Извлекает числовые значения из строк (поддерживает как целые числа, так и дробные).

6. **Функция `magic`**
   - Основная функция, которая обрабатывает пользовательский ввод через консоль и генерирует результаты для выбранного класса и четверти.

7. **Графический интерфейс с `tkinter`**
   - В проекте также предусмотрен графический интерфейс с использованием библиотеки `tkinter`. При его использовании пользователь может вводить данные через графические поля ввода вместо консольного ввода.

## Установка и настройка

Для использования этого проекта вам потребуется установить следующие зависимости:

```bash
pip install requests pandas openpyxl beautifulsoup4 lxml