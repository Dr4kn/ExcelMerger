import pandas
import pandas as pd


# Separate function to easily comment out the calling code
def insert_columns(merge_into: pandas.DataFrame, merge_from: pandas.DataFrame, at_index: int, array: [str]):
    for item in array:
        merge_into.insert(at_index, item, merge_from[item])
        at_index += 1
    return merge_into


def merge_old_into_new(old_file: str, new_file: str):
    # Load the first Excel file into a pandas DataFrame, skipping the first row
    merge_into = pd.read_excel(old_file, skiprows=1)

    # Load the second Excel file into a pandas DataFrame, skipping the first row
    new = pd.read_excel(new_file)

    # Only get first part of the value
    merge_into['X-nat Pseudonym'] = merge_into['X-nat Pseudonym'].str.split('_').str.get(0)

    # Convert "Datum MRT" column to datetime format with different date formats
    iso8601 = '%Y-%m-%d'

    # Date Format: DD/MM/YYYY
    merge_into['Datum MRT'] = pd.to_datetime(merge_into['Datum MRT'], dayfirst=True).dt.strftime(iso8601)
    # Date Format: MM.DD.YYYY
    # the worst data format is used ...
    new['Datum MRT'] = pd.to_datetime(new['Datum MRT'], format='%m/%d/%Y').dt.strftime(iso8601)

    # Merge the two DataFrames based on "X-nat Pseudonym" and "Datum MRT"
    old = pd.merge(merge_into, new, on=['X-nat Pseudonym', 'Datum MRT'], how='left')

    # Columns in new to insert into merge_into
    columns_to_merge = ['scan_description', 'scan_id', 'experiment_id', 'Exclude->no whole body', 'comment informatics',
                        'comment radiology']

    # Insert merged data
    old = insert_columns(new, old, 9, columns_to_merge)

    # Write the merged DataFrame to a new Excel file
    old.to_excel('merged.xlsx', index=False)


if __name__ == '__main__':
    merge_old_into_new("old.xlsx", "new.xlsx")
    print("done")
