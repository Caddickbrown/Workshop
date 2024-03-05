import csv
import os
from datetime import datetime

def create_csv_if_not_exists(filename):
    try:
        if not os.path.exists(filename):
            with open(filename, 'w', newline='') as csvfile:
                fieldnames = ['Date', 'Info']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                print(f"CSV file '{filename}' created successfully.")
        else:
            print(f"CSV file '{filename}' already exists.")
    except Exception as e:
        print(f"An error occurred: {e}")

def append_to_csv(filename, date, info):
    # Check if the file exists, create if it doesn't
    create_csv_if_not_exists(filename)

    # Open the CSV file in append mode
    with open(filename, 'a', newline='') as csvfile:
        fieldnames = ['Date', 'Info']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        # Write the new row
        writer.writerow({'Date': date, 'Info': info})

def main():
    filename = 'data.csv'

    # Get date and info from the user
    date = input("Enter the date (YYYY-MM-DD) [Leave blank for today's date]: ")
    if date == '':
        date = datetime.now().strftime('%Y-%m-%d')

    info = input("Enter the info: ")

    # Append to CSV
    append_to_csv(filename, date, info)

    print("Data appended to", filename)

if __name__ == "__main__":
    main()
