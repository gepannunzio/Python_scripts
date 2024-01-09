import os
import time
from pathlib import Path
from zipfile import ZipFile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

# Navigate to a page and click on an element
def navigate_and_click(driver, url, locator, timeout=10):
    driver.get(url)
    element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(locator))
    element.click()

# Choose options from dropdowns
def choose_options(driver, dropdowns, options):
    for dropdown_key, dropdown_id in dropdowns.items():
        dropdown = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, dropdown_id)))
        dropdown.click()
        option = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, options[dropdown_key])))
        option.click()

# Wait for file download completion
def wait_for_download(downloads_directory, filename, timeout=60):
    filepath = os.path.join(downloads_directory, filename)
    start_time = time.time()
    while not os.path.exists(filepath) or time.time() - start_time > timeout:
        time.sleep(1)

# Main function to execute the entire process
def main():
    driver_path = './msedgedriver'
    driver = webdriver.Chrome()

    # Navigate to the "Documents & Downloads" section
    url = "https://cde.ucr.cjis.gov/LATEST/webapp/#/pages/home"
    navigate_and_click(driver, url, (By.LINK_TEXT, 'Documents & Downloads'))

    # Set dropdowns for downloading parameters
    dropdowns = {'table': 'dwnnibrs-download-select', 'year': 'dwnnibrscol-year-select', 'location': 'dwnnibrsloc-select'}
    options = {'table': '/html/body/ngx-app/ngx-pages/ngx-one-column-layout/nb-layout/div[2]/div/div/nb-option-list/ul/nb-option[4]',
               'year': '/html/body/ngx-app/ngx-pages/ngx-one-column-layout/nb-layout/div[2]/div/div/nb-option-list/ul/nb-option[1]',
               'location': '/html/body/ngx-app/ngx-pages/ngx-one-column-layout/nb-layout/div[2]/div/div/nb-option-list/ul/nb-option[10]'}
    choose_options(driver, dropdowns, options)

    time.sleep(5)

    download_button_id = "nibrs-download-button"
    download_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, download_button_id)))
    download_button.click()

    # Local Variables for downloading and extracting zip file

    # Detect the user's Downloads directory
    downloads_directory = str(Path.home() / "Downloads")
    wait_for_download(downloads_directory, "victims.zip")

    zip_filename = os.path.join(downloads_directory, "victims.zip")
    dest_folder = './'
    excel_filename = "Victims_Age_by_Offense_Category_2022.xlsx"

    # Extract the contents of the ZIP file
    with ZipFile(zip_filename, 'r') as zip_ref:
        zip_ref.extract(excel_filename, dest_folder)

    # Load the Excel file using pandas
    df = pd.read_excel(excel_filename, sheet_name="Table 5 NIBRS 2022", engine="openpyxl", index_col=0)

    if 'Crimes Against Property' in df.index:
        # Filter for 'Crimes Against Property totals'
        filtered_df = df.loc['Crimes Against Property']

        # Drop the index column and generate CSV without the index
        filtered_df.reset_index(drop=True, inplace=True)

        # Save the filtered data to a CSV file
        csv_filename = "Crimes_Against_Property_2022_Totals.csv"
        filtered_df.to_csv(csv_filename, index=False)

        # Extract details related to 'Crimes Against Property'
        details_df = df.loc['Crimes Against Property':].copy()

        # Remove the column 'Total Victims' by index
        details_df = details_df.drop(details_df.columns[0], axis=1, errors='ignore')
        
        # Remove the first row (Crimes Against Property totals). Comment this line if you want to see the sum of the details per age
        # Remove the last row (footer)
        details_df = details_df.iloc[1:-1]

        # Drop the index column and generate CSV without the index
        details_df.reset_index(drop=True, inplace=True)

        # Save the details to a separate CSV file
        details_csv_filename = "Crimes_Against_Property_Details_2022.csv"
        details_df.to_csv(details_csv_filename, index=False, header=False)

        # Print or further process the details DataFrame
        print("\nDetails DataFrame (Crimes Against Property):")
        print(details_df)
    else:
        print("Row 'Crimes Against Property' not found in the DataFrame.")


    # Clean up: Remove downloaded files
    os.remove(zip_filename)
    os.remove(excel_filename)

    print(f"CSV file '{csv_filename}' has been generated successfully.")

    # Close the WebDriver
    driver.quit()

# Run the main function
if __name__ == "__main__":
    main()