import os
import pandas as pd

# Directory containing the Excel files
directory = os.path.join(os.getcwd(), 'BC History')

# Dictionary to store the merged data
merged_data = {}

# Get the list of Excel files in the directory
excel_files = [filename for filename in os.listdir(directory) if filename.endswith('.xlsx')]

# Sort the Excel files by start number in descending order
excel_files.sort(key=lambda x: int(x.split('_')[0]), reverse=True)
print("\nfile name :: number of Data")
# Iterate over the sorted Excel files
for filename in excel_files:
    # Extract start and end numbers from the file name
    start, end = map(int, os.path.splitext(filename)[0].split('_'))

    # Read the Excel file into a DataFrame, skipping the header row
    filepath = os.path.join(directory, filename)
    df = pd.read_excel(filepath, header=None)

    print(f"{filename} :: {len(df.index)}")
    # Store the Hash Content and Bang Content for each row in the dictionary
    for index, row in df.iterrows():
        row_no = row[0]  # Assuming the row numbers are in the first column
        hash_content = row[1]
        bang_content = row[2]

        merged_data[row_no] = {
            'Hash Content': hash_content,
            'Bang Content': bang_content
        }

merged_data_length = len(merged_data)
print(f"\nThe length of the merged data is: {merged_data_length}")

partCount = 0
critical = 1.5

# Iterate over the merged data dictionary
for row_no, content in merged_data.items():
    try:
        bang_content = float(content['Bang Content'])
        # print(bang_content)
        if bang_content < critical:
           partCount = partCount + 1
        else:
            if partCount > 8:
                delta = row_no + partCount
                print(f"\n************ Bang < {critical} => Count : {partCount} | From ~ To : {row_no + 1} ~ {row_no + partCount} ************")
                for ii in range(row_no + 1, row_no + partCount + 1):
                    print(ii, merged_data[ii])
            partCount = 0
    except ValueError:
        print("\nSkip in Header : ",row_no)
        continue  # Skip this row if the conversion fails
 

   