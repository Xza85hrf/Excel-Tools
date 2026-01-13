# Changelog

## [Unreleased] - 2025-04-25

### Added

- Core Excel file comparison functionality (Exact and Contains strategies).
- Graphical User Interface (GUI) using Tkinter.
- Highlighting of found (green) and missing (red) rows in the source Excel file.
- Option to include subdirectories in the search.
- Export results table to CSV file functionality.
- Basic application logging (viewable via File -> View Log).
- Full Polish language translation for the user interface.
- Configuration persistence: Saves last used paths, column, extension, strategy, and subdirectory settings to `config.json`.
- Detailed, formatted help/instructions window accessible from the menu.
- Application icon (`icons8.ico`).
- `requirements.txt` for managing dependencies.
- `README.md` with detailed usage and build instructions.

### Changed

- Improved help window formatting for better readability.
- Made help window resizable and increased default font size.
- Optimized executable build process using a virtual environment to reduce file size.

### Removed

- Non-functional 'startswith' and 'regex' matching strategies (temporarily, can be added back).
