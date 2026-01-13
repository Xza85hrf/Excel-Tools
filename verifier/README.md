# Excel File Verifier

Compare filenames listed in an Excel sheet against files present in a specified directory and its subdirectories. Highlights found/missing rows in the Excel file and provides a results table.

## Features (Funkcje)

* Specify the Excel file containing the list of filenames.
* Specify the directory to search within.
* Choose the column in the Excel file that contains the filenames.
* Specify the file extension to look for (e.g., `.pdf`, `.docx`).
* Choose the matching strategy:
  * **Dokładna (Exact):** Finds files with names that exactly match the entry in the Excel sheet (case-insensitive).
  * **Zawiera (Contains):** Finds files where the name contains the entry from the Excel sheet (case-insensitive).
* Optionally include subdirectories in the search.
* Displays results directly in the application window, indicating found/not found status.
* Highlights corresponding rows in the Excel file (Green for found, Red for not found).
* **NEW:** Export results to a CSV file.
* **NEW:** User interface fully translated into Polish.
* **NEW:** Saves your last used settings (paths, column, extension, strategy, subdirectory option) to `config.json` for convenience.

## Requirements (Wymagania)

* Python 3.x
* Required libraries (zainstaluj używając `pip install -r requirements.txt`):
  * `openpyxl >= 3.0.0`
  * `rich >= 12.0.0`
  * `pyinstaller >= 4.0.0` (for building the executable)

## How to Run (Jak Uruchomić)

1. **Clone or download** this repository.
2. **Navigate** to the application directory in your terminal:

    ```bash
    cd /path/to/Excel-File-Verifier
    ```

3. **Install** the required libraries:

    ```bash
    pip install -r requirements.txt
    ```

4. **Run** the application:

    ```bash
    python excel_verifier_gui.py
    ```

## Using the Application (Używanie Aplikacji)

1. **Uruchom aplikację** (`python excel_verifier_gui.py`).
2. **Wybierz plik Excel:** Kliknij "Przeglądaj..." obok "Plik Excel" i wybierz plik `.xlsx`.
3. **Wybierz katalog:** Kliknij "Przeglądaj..." obok "Katalog" i wybierz folder do przeszukania.
4. **Wpisz nazwę kolumny:** W polu "Nazwa Kolumny" wpisz literę kolumny (np. `A`, `B`) zawierającej nazwy plików.
5. **Wpisz rozszerzenie pliku:** W polu "Rozszerzenie Pliku" wpisz rozszerzenie (np. `.pdf`, `.docx`) szukanych plików.
6. **Wybierz strategię dopasowania:** Wybierz "Dokładna" lub "Zawiera".
7. **Uwzględnij podkatalogi:** Zaznacz pole, jeśli chcesz przeszukiwać również podkatalogi.
8. **Kliknij "Sprawdź Pliki".**
9. **Wyniki:** Aplikacja wyświetli wyniki w dolnym panelu i zaktualizuje plik Excel kolorami. Możesz również kliknąć "Eksportuj do CSV", aby zapisać wyniki.

## Building the Executable (Tworzenie Pliku Wykonywalnego)

You can create a standalone `.exe` file that can be run on Windows without needing Python installed.

1. Make sure you have installed the requirements (`pip install -r requirements.txt`).
2. Navigate to the application directory in your terminal.
3. Run the following command:

    ```bash
    pyinstaller --name "Excel File Verifier" --onefile --noconsole --icon icons8.ico excel_verifier_gui.py
    ```

    * `--name "Excel File Verifier"`: Sets the name of the output executable (and spec file).
    * `--onefile`: Creates a single executable file.
    * `--noconsole`: Prevents the command prompt window from appearing when the application runs.
    * `--icon icons8.ico`: Specifies the application icon (make sure `icons8.ico` is in the same directory).
4. The executable file (`Excel File Verifier.exe`) will be created in the `dist` subfolder.

## Troubleshooting (Rozwiązywanie Problemów)

* **Error reading Excel file:** Ensure the file path is correct and the file is not corrupted. Make sure the specified column exists.
* **Permission errors:** Ensure the application has permission to read the specified directory and write to the Excel file.
* **Incorrect results:** Double-check the column name, file extension, and matching strategy. Ensure filenames in Excel don't have hidden characters or leading/trailing spaces.
