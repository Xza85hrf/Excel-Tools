# Future Development Plans

This document outlines potential enhancements and features for future versions of the Excel Comparison Application.

## Potential Features

* **Advanced Matching Strategies:**
  * Re-implement and refine 'Starts With' strategy.
  * Implement 'Regular Expression (Regex)' strategy for complex pattern matching.
  * Explore fuzzy matching (e.g., Levenshtein distance) for better suggestions on missing files.
* **User Interface Enhancements:**
  * Add user controls for adjusting font sizes (e.g., in the help window or main UI).
  * Implement a visual progress bar for long-running comparison operations.
  * Improve the layout or look-and-feel of the GUI.
  * Add options to customize the highlighting colors.
* **Core Functionality:**
  * Support for older `.xls` Excel file formats (might require `xlrd`).
  * Option to make matching case-sensitive.
  * Allow specifying multiple columns in the Excel file to check.
  * Option to specify multiple file extensions simultaneously (e.g., `.pdf, .docx`).
  * Handle potential errors more gracefully (e.g., file permission issues, corrupted Excel files) with more specific user feedback.
* **Quality & Maintainability:**
  * Add unit tests and integration tests to ensure reliability.
  * Refactor code for improved clarity or performance.
  * Enhance logging with more detail or different log levels.
* **Build & Distribution:**
  * Explore alternative build tools or methods if further size reduction is needed.
  * Consider creating installers for easier distribution.
