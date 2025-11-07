# 🧩 Kundoluk Parser

A tool for processing and analyzing student performance data in **Excel** format.  
It uses the `tkinter` library to create a graphical user interface (GUI) and supports multithreading to perform long operations without freezing the interface.  
The program calculates averages, categorizes students, and generates visual charts 📊.

---

## ✨ Features

- **📥 Data Input**  
  The user enters the class (e.g., `4B`) and the term number to process the corresponding data.  

- **⏳ Preloader**  
  A visual progress indicator shows that the data is being processed.  

- **🧮 Data Processing**
  - Analyzes data from Excel files.  
  - Calculates each student's average score.  
  - Categorizes students into:  
    **Excellent**, **Good**, **Satisfactory**, **Poor**, **Unsatisfactory**.  
  - Generates a pie chart with category distribution 🎯.  

- **💻 Graphical Interface**
  - Simple window for data input.  
  - Button to start data processing.  
  - Stylish design: buttons and input fields change color when hovered 🌈.  

---

## 🧱 Project Structure

### **KParser** – class for interacting with the Kundoluk system  
**Main Functions:**  
- `get_grade()` – retrieves class data.  
- `get_subject()` – retrieves subject data.  
- `magic()` – handles user input.  

---

### **Grade** – creates an Excel file with subject data  
**Main Functions:**  
- Generates an Excel table based on data fetched from the Kundoluk system.  

---

### **assign_categories** – adds performance categories  
**Arguments:**  
- `result_sheet` – Excel sheet with results.  
- `num_subjects` – number of subjects.  

**Features:**  
- Adds performance labels (Excellent, Good, etc.) next to the average score.  

---

### **calculate_averages** – calculates averages and performance summary  
**Capabilities:**  
- Creates a **Result** sheet containing:  
  - Average scores per subject and overall average.  
  - Performance labels based on the final average.  
- Generates a small summary table with percentages and a chart 📈.  

---

### **start_ui** – creates the graphical interface  
**Highlights:**  
- Includes a preloader showing file creation progress.  
- All windows are centered automatically for better usability 🎯.  

---

## ⚠️ Important!
- **Before launching**, log in to the Kundoluk system on the device where the program will run.  
- Ensure cookies are enabled in your browser.  

---

## ⚙️ Installation

1. Download the ZIP file.  
2. Install dependencies:  
   ```bash
   pip install -r requirements.txt
   ```

## Preview:
![previewimage](https://github.com/Lollypop-stack/openpyxlCompilerForSchool/blob/main/openpyxlCompilerForSchool-main/AppPreview/%D0%A1%D0%BD%D0%B8%D0%BC%D0%BE%D0%BA%20%D1%8D%D0%BA%D1%80%D0%B0%D0%BD%D0%B0%202025-11-07%20151515.png)
#
![previewimage](https://github.com/Lollypop-stack/openpyxlCompilerForSchool/blob/main/openpyxlCompilerForSchool-main/AppPreview/%D0%A1%D0%BD%D0%B8%D0%BC%D0%BE%D0%BA%20%D1%8D%D0%BA%D1%80%D0%B0%D0%BD%D0%B0%202025-11-07%20153144.png)
#

# 🧩 Kundoluk Parser

Инструмент для обработки и анализа данных об успеваемости учеников в формате **Excel**.  
Использует библиотеку `tkinter` для создания удобного графического интерфейса и потоки для многозадачности.  
Программа вычисляет средние баллы, распределяет учеников по категориям и создаёт диаграммы 📊.

---

## ✨ Основные функции

- **📥 Ввод данных**  
  Пользователь вводит класс (например, `4Б`) и номер четверти для обработки данных.  

- **⏳ Прелоадер**  
  Визуальный индикатор показывает процесс обработки данных.  

- **🧮 Обработка данных**
  - Анализирует Excel-файлы.  
  - Считает средний балл каждого ученика.  
  - Делит учеников на категории: **Отлично**, **Хорошо**, **Удовлетворительно**, **Плохо**, **Неудовлетворительно**.  
  - Создаёт круговую диаграмму с распределением 🎯.  

- **💻 Графический интерфейс**
  - Простое окно для ввода.  
  - Кнопка запуска обработки.  
  - Современный дизайн с динамичными элементами: кнопки и поля меняют внешний вид при наведении 🌈.  

---

## 🧱 Структура проекта

### **KParser** – класс для работы с системой Kundoluk  
**Основные функции:**  
- `get_grade()` – получение данных о классе.  
- `get_subject()` – получение данных о предметах.  
- `magic()` – обработка пользовательского ввода.  

---

### **Grade** – создание Excel-файла с предметными данными  
**Основные функции:**  
- Создаёт таблицу Excel на основе информации, полученной со страницы класса в системе Kundoluk.  

---

### **assign_categories** – добавление категорий успеваемости  
**Аргументы:**  
- `result_sheet` – лист Excel с результатами.  
- `num_subjects` – количество предметов.  

**Функции:**  
- Добавляет категории (Отл., Хор., Уд., Пл., Нез.) рядом со средним баллом.  

---

### **calculate_averages** – расчёт средних значений и итогов  
**Возможности:**  
- Создаёт лист **Result**, где:  
  - Подсчитываются средние баллы по предметам и общий итог.  
  - Добавляется уровень успеваемости на основе итогового среднего.  
- Формирует мини-таблицу с процентным распределением успеваемости класса и строит диаграмму 📈.  

---

### **start_ui** – создание графического интерфейса  
**Особенности:**  
- Включён прелоадер, показывающий процесс создания файла.  
- Все окна автоматически центрируются на экране для удобства использования 🎯.  

---

## ⚠️ Важно!

- **Перед запуском** войдите в систему Kundoluk на том устройстве, где планируете запускать программу.  
- Убедитесь, что в вашем браузере разрешены Cookies.

---

## ⚙️ Установка

1. Скачайте ZIP-архив проекта.  
2. Установите зависимости:  
   ```bash
   pip install -r requirements.txt
   ```
#
## Превью:
![previewimage](https://github.com/Lollypop-stack/openpyxlCompilerForSchool/blob/main/openpyxlCompilerForSchool-main/AppPreview/%D0%A1%D0%BD%D0%B8%D0%BC%D0%BE%D0%BA%20%D1%8D%D0%BA%D1%80%D0%B0%D0%BD%D0%B0%202025-11-07%20151515.png)
#
![previewimage](https://github.com/Lollypop-stack/openpyxlCompilerForSchool/blob/main/openpyxlCompilerForSchool-main/AppPreview/%D0%A1%D0%BD%D0%B8%D0%BC%D0%BE%D0%BA%20%D1%8D%D0%BA%D1%80%D0%B0%D0%BD%D0%B0%202025-11-07%20153144.png)
#
