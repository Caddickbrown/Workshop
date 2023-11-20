# This is a work in progress and doesn't seem to work (Not sure why yet)

import csv
import tkinter as tk
from tkinter import filedialog
import os

def csv_to_rec(csv_filename):
    # Get the absolute path of the CSV file
    csv_path = os.path.abspath(csv_filename)

    # Create the REC file path by changing the extension of the CSV file
    rec_path = os.path.splitext(csv_path)[0] + ".rec"

    try:
        # Read the CSV file and store rows in a list
        with open(csv_path, 'r') as csv_file:
            csv_reader = csv.DictReader(csv_file)
            rows = list(csv_reader)

        # Create the directory structure leading to the REC file
        os.makedirs(os.path.dirname(rec_path), exist_ok=True)

        try:
            # Open the REC file with 'x' mode to create a new file
            with open(rec_path, 'x') as rec_file:
                # Write each row to the REC file
                for row in rows:
                    rec_file.write("# Record\n")
                    for key, value in row.items():
                        rec_file.write(f"{key}: {value}\n")
                    rec_file.write("\n")

            print(f"Conversion complete. Recfile saved as {rec_path}")
        except FileExistsError:
            print(f"Error: REC file already exists at {rec_path}")
    except Exception as e:
        print(f"An error occurred during conversion: {e}")

    return rec_path  # Return the generated recfile path

def select_csv_file():
    # Create a Tkinter root window (it won't be shown)
    root = tk.Tk()
    # Hide the main window
    root.withdraw()

    # Open a file dialog to select a CSV file
    file_path = filedialog.askopenfilename(
        title="Select CSV file",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )

    # Check if a valid CSV file is selected
    if file_path and file_path.lower().endswith(".csv"):
        return file_path
    else:
        print("No or invalid CSV file selected.")
        return None

if __name__ == "__main__":
    # Select a CSV file using the file dialog
    csv_input = select_csv_file()

    # If a valid CSV file is selected, convert it to REC format
    if csv_input:
        csv_to_rec(csv_input)

    # Wait for user input before exiting
    input("Press Enter to exit...")