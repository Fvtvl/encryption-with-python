import os
import csv
from spire.xls import *


folder_path = "Excels"
csv_passwords_file = "Eduard_Haponov - passwords_1.csv"
csv_list_file_name = 'Eduard_Haponov - Passwords list.csv'
excel_file_names = []
passwords = []


def get_excel_names(folder_path):
    try:
        excel_files = []
        for file_name in os.listdir(folder_path):
            if file_name.endswith('.xlsx'):
                excel_files.append(os.path.join(folder_path, file_name))
        return excel_files
    except Exception as e:
        print(f"An error occurred while getting Excel file names: {e}")
        return []


def load_passwords_from_csv(csv_passwords_file):
    try:
        passwords = []
        with open(csv_passwords_file, 'r', newline='') as file:
            reader = csv.reader(file)
            next(reader)
            for row in reader:
                password = row[0]
                passwords.append(password)
        return passwords
    except Exception as e:
        print(f"An error occurred while loading passwords from CSV: {e}")
        return []


def create_csv_dict(excel_file_names, passwords):
    try:
        file_password_pairs = zip(excel_file_names, passwords)
        csv_dict = {file: password for file, password in file_password_pairs}
        return csv_dict
    except Exception as e:
        print(f"An error occurred while creating CSV dictionary: {e}")
        return {}


def encrypt_excel_files(csv_dict):
    try:
        for excel_file, password in csv_dict.items():
            workbook = Workbook()
            workbook.LoadFromFile(excel_file)
            workbook.Protect(password)
            workbook.SaveToFile(f"Protected_{excel_file}", ExcelVersion.Version2013)
            workbook.Dispose()
        print("All Excel files have been successfully encrypted.")
    except Exception as e:
        print(f"An error occurred while encrypting Excel files: {e}")


def create_csv(data, csv_list_filename):
    try:
        with open(csv_list_filename, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['File name', 'Password'])
            for file_name, password in data.items():
                writer.writerow([file_name, password])
        print(f"CSV file '{csv_list_filename}' has been successfully created.")
    except Exception as e:
        print(f"An error occurred while creating CSV file: {e}")



excel_file_names = get_excel_names(folder_path)
passwords = load_passwords_from_csv(csv_passwords_file)
csv_dict = create_csv_dict(excel_file_names, passwords)

encrypt_excel_files(csv_dict)
create_csv(csv_dict, csv_list_file_name)

