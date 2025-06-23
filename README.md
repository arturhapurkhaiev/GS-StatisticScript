* * * I use ChatGPT to create this script * * * 
 * 📊 Google Sheets Script: Sheet and Column Statistics with Bilingual Support
 *
 * 🇬🇧 English Description:
 * This script adds a custom menu to Google Sheets that lets users calculate and display
 * statistical information about their dataset from a sheet named "Data". Results are written
 * into a separate sheet named "Statistic".
 *
 * ✨ Key Features:
 * - Adds a custom menu: "Statistics" → "Calculate statistics"
 * - Analyzes the "Data" sheet and outputs a statistical summary to "Statistic"
 * - Automatically creates tables for both sheet-level and column-level metrics
 * - Applies formatting and borders for readability
 * - Supports both Ukrainian 🇺🇦 and English 🇬🇧 languages
 *
 * 📌 Sheet-level Statistics:
 * - Number of rows with data
 * - Number of completely empty rows
 * - Number of columns
 * - Number of unique rows
 * - Number of duplicate rows
 *
 * 📌 Column-level Statistics:
 * - Column name (from the header row)
 * - Number of empty cells
 * - Number of unique values
 * - Average (for numeric columns)
 * - Median (for numeric columns)
 *
 * 🌐 Localization:
 * The script supports **bilingual output**: Ukrainian and English.
 * - Language is controlled by the `LANGUAGE` constant at the top of the script.
 * - To switch language, simply change:
 *     const LANGUAGE = 'ua'; // for Ukrainian
 *     const LANGUAGE = 'en'; // for English
 *
 * 🛠 Requirements:
 * - Sheets named exactly "Data" and "Statistic" must exist before running
 * - The first row in "Data" should contain column headers
 *
 * 🧠 Ideal For:
 * - Quick data profiling
 * - Identifying missing or duplicate data
 * - Summarizing datasets for reports or cleaning
                                ** 
                  --Open your Google Sheet.
                  --From the menu, select Extensions → Apps Script.
                  --In the script editor window that opens, delete any existing code (if needed).
                  --Paste the provided script there.
                  --Save the project (Ctrl + S or click the floppy disk icon).
                  --Close the script editor.
                  --Refresh the Google Sheet page to see the new menu.
                  --Use the new menu item to run the statistics function.
 * 🇺🇦 Опис українською:
 * Цей скрипт додає власне меню до Google Sheets і дозволяє обчислювати та виводити
 * статистичну інформацію з аркуша з назвою "Data". Результати виводяться в аркуш "Statistic".
 *
 * ✨ Основні можливості:
 * - Додає меню: "Статистика" → "Обчислити статистику"
 * - Аналізує дані з аркуша "Data" і виводить зведену статистику в "Statistic"
 * - Створює таблиці для статистики по всьому аркушу та по кожній колонці
 * - Форматує та додає межі для кращої читабельності
 * - Підтримує українську 🇺🇦 та англійську 🇬🇧 мови
 *
 * 📌 Статистика по аркушу:
 * - Кількість рядків із даними
 * - Кількість повністю порожніх рядків
 * - Кількість стовпців
 * - Кількість унікальних рядків
 * - Кількість дублікатів
 *
 * 📌 Статистика по колонках:
 * - Назва колонки (з заголовка першого рядка)
 * - Кількість порожніх клітинок
 * - Кількість унікальних значень
 * - Середнє значення (для числових колонок)
 * - Медіана (для числових колонок)
 *
 * 🌐 Локалізація:
 * Скрипт підтримує двомовний інтерфейс: українською та англійською.
 * - Мова визначається за допомогою константи `LANGUAGE` на початку скрипта:
 *     const LANGUAGE = 'ua'; // українська
 *     const LANGUAGE = 'en'; // англійська
 *
 * 🛠 Вимоги:
 * - У таблиці мають бути аркуші з назвами "Data" і "Statistic"
 * - Перший рядок у "Data" повинен містити заголовки колонок
 *
 * 🧠 Ідеально підходить для:
 * - Швидкого аналізу даних
 * - Виявлення порожніх або дублікатних рядків
 * - Створення короткої аналітики для звітів або очищення даних
                         ***
                  --Відкрийте вашу Google Таблицю.
                  --У меню виберіть Розширення (Extensions) → Apps Script.
                  --У відкритому вікні редактора скриптів видаліть увесь існуючий код (якщо потрібно).
                  --Вставте туди скрипт, який я надав.
                  --Збережіть проект (Ctrl + S або іконка дискети).
                  --Закрийте редактор скриптів.
                  --Оновіть сторінку Google Таблиці, щоб меню з’явилося.
                  --Використовуйте нове меню для запуску функції статистики.
