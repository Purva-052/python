import re
import pandas as pd
from google_play_scraper import app 

def add_status_to_excel(input_file):
    try:
        # Ensure the input_file has the .xlsx extension
        if not input_file.endswith('.xlsx'):
            input_file += '.xlsx'

        # Read input Excel file
        df = pd.read_excel(input_file, header=None)  # Read without header

        status_column = []

        for row in df.itertuples(index=False):
            url = row[0]
            if isinstance(url, str):
                matches = re.findall(r'https://play\.google\.com/store/apps/details\?id=([\w\.]+)', url)

                if matches:
                    package_name = matches[0].strip()
                    status = "open"

                    try:
                        app_info = app(package_name)
                    except Exception as e:
                        status = "close"
                else:
                    package_name = ""
                    status = ""
            else:
                package_name = ""
                status = ""

            status_column.append(status)

        # Add the status information to the original Excel file
        df['Status'] = status_column
        df.to_excel(input_file, index=False, header=False)

        print(f"Status added to the '{input_file}' file.")

    except FileNotFoundError:
        print("Input Excel file not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    input_file = input("Enter the name of the input Excel file (without extension .xlsx): ")
    add_status_to_excel(input_file)
