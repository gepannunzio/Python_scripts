import pandas as pd

# Provide the correct Excel file path and sheet name

# Read the Excel file into a DataFrame
excel_filename = "Victims_Age_by_Offense_Category_2022.xlsx"
df = pd.read_excel(excel_filename, sheet_name="Table 5 NIBRS 2022", engine="openpyxl")


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
    details_df = details_df.iloc[1:]

    # Remove the last row (footer)
    details_df = details_df.iloc[:-1]

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
