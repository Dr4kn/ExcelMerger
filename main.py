import pandas as pd


def merge_excel_files(old_file, new_file):
    # Load the first Excel file into a pandas DataFrame, skipping the first row
    df1 = pd.read_excel(old_file, skiprows=1)

    # Load the second Excel file into a pandas DataFrame, skipping the first row
    df2 = pd.read_excel(new_file)

    df1['X-nat Pseudonym'] = df1['X-nat Pseudonym'].str.split('_').str.get(0)
    # Convert "Datum MRT" column to datetime format with different date formats
    df1['Datum MRT'] = pd.to_datetime(df1['Datum MRT'], dayfirst=True)
    df2['Datum MRT'] = pd.to_datetime(df2['Datum MRT'], format='%m/%d/%Y')

    # Merge the two DataFrames based on "X-nat Pseudonym" and "Datum MRT"
    merged_df = pd.merge(df1, df2, on=['X-nat Pseudonym', 'Datum MRT'], how='left')

    # Write the merged DataFrame to a new Excel file
    merged_df.to_excel('merged.xlsx', index=False)


if __name__ == '__main__':
    merge_excel_files("old.xlsx", "new.xlsx")
    print("done")
