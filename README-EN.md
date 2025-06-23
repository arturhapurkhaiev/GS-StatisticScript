# GS-StatisticScript
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
