Excel Random Data Generator

Excel Random Data Generator is a Python application built with the Tkinter, Faker, pandas and openpyxl libraries. It allows users to generate random client data and save it to Excel files.

Features

Generate custom number of client records.
Specify custom headers for the data.
Generate random client names and addresses using the Faker library.
Store data in Excel files.

Installation

Prerequisites
Python 3.x
pip

Install Dependencies

Clone this repository to your local system and navigate to the cloned directory. Then, install the required dependencies using pip:


pip install -r requirements.txt

Usage

To start the application, navigate to the project directory and run:

python main.py

In the GUI:

Enter the headers for the data you want to generate, separated by commas.
Enter the number of clients you want to generate for up to four sheets.
Click 'Generate and save' to generate the data and save it to Excel files.
You can also click 'Generate Default' to quickly generate default data for 20 and 30 clients.

Contributing
Contributions are welcome! Please feel free to submit a Pull Request or open an issue.
