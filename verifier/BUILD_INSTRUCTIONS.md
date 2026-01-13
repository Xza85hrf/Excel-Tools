# Building the Executable (Optimized)

This guide explains how to build the standalone `.exe` file for the **Excel File Verifier** application using PyInstaller within a clean virtual environment. Using a virtual environment is highly recommended as it significantly reduces the final executable size by only including necessary dependencies.

## Steps

1.  **Navigate to the Project Directory:**
    Open your terminal (like PowerShell or Command Prompt) and navigate to the application's root directory:
    ```bash
    cd c:\Users\Barbatos\Desktop\Excel-File-Verifier
    ```

2.  **Create a Virtual Environment (if it doesn't exist):**
    If you haven't already created a virtual environment named `.venv` within the project folder, run:
    ```bash
    python -m venv .venv
    ```
    This creates a folder named `.venv` containing a clean Python installation isolated from your system's main Python.

3.  **Activate the Virtual Environment:**
    You need to activate the environment to use it. The command differs slightly depending on your shell:

    *   **PowerShell:**
        ```powershell
        .\.venv\Scripts\Activate.ps1
        ```
        *Note: If you encounter an error about script execution being disabled, you might need to temporarily bypass the policy for the activation command. See step 4.* 

    *   **Command Prompt (cmd.exe):**
        ```bash
        .\.venv\Scripts\activate.bat
        ```

    Your terminal prompt should change, usually prefixed with `(.venv)`, indicating the virtual environment is active.

4.  **Install Dependencies:**
    Once the virtual environment is active, install the required packages listed in `requirements.txt`:
    ```bash
    pip install -r requirements.txt
    ```
    *(If activating the environment directly failed in PowerShell due to execution policy, you can combine activation and installation like this):*
    ```powershell
    powershell -ExecutionPolicy Bypass -Command ".\.venv\Scripts\Activate.ps1; pip install -r requirements.txt"
    ```

5.  **Build the Executable:**
    With the virtual environment still active and dependencies installed, run the PyInstaller command:
    ```bash
    pyinstaller --name "Excel File Verifier" --onefile --noconsole --icon icons8.ico excel_verifier_gui.py
    ```
    *(Alternatively, if activating directly failed, combine activation and build):*
    ```powershell
    powershell -ExecutionPolicy Bypass -Command ".\.venv\Scripts\Activate.ps1; pyinstaller --name \"Excel File Verifier\" --onefile --noconsole --icon icons8.ico excel_verifier_gui.py"
    ```

    *   `--name "Excel File Verifier"`: Sets the output executable name.
    *   `--onefile`: Bundles everything into a single `.exe` file.
    *   `--noconsole`: Prevents a console window from appearing when the GUI app runs.
    *   `--icon icons8.ico`: Sets the application icon.
    *   `excel_verifier_gui.py`: The main script file for the application.

6.  **Find the Executable:**
    PyInstaller will create `build` and `dist` folders. Your final executable, `Excel File Verifier.exe`, will be located inside the `dist` folder.

7.  **Deactivate the Virtual Environment (Optional):**
    When you are finished building, you can deactivate the virtual environment by simply running:
    ```bash
    deactivate
    ```
    Your terminal prompt will return to normal.

Following these steps ensures you build the smallest possible executable by isolating the build process from potentially unnecessary packages in your global Python environment.

