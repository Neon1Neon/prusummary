import pandas as pd
import re
from datetime import datetime
import io # You'll need this to read the uploaded file content into pandas

# Define the correct column names for the tax data
tax_data_columns = [
    '1ST PUB', '2ND PUB', 'FORFEITED PUB', 'REPUB', 'ERRATUM',
    'VOLUNTARY WITHDRAWAL', 'EXPIRED (INV/UM/ID)', 'VERIFIED EXPIRED PATENTS',
    'NOTICE OF 1ST PUBLICATION', 'VERIFIED 2ND PUB & ISSUANCE FEE',
    'LIST OF GRANT/ REGISTRATION', 'E-CERT /NOTICES', 'REVIEWED CERTIFICATES',
    'ECORR CERTIFICATE', 'MAILED CERTIFICATE', 'RELEASED CERT',
    'CERT OF CORRECTION', 'RE-ISSUANCE OF CERT', 'REMINDER (INV/ID)',
    'ENCODING OF ANNUITY DETAILS', 'E-DOCKETING', 'ANNUITY DEFICITS (INV/ID)',
    'NNP', 'VERIFIED PROCESSED ANNUITIES', 'LAPSED', 'RECORDAL',
    'PUB-RECORDAL', 'ID RENEWALS', 'PUB-RENEWAL', 'REVIEWED ID RENEWALS',
    'EMAIL RESPONSE', 'REQUEST FOR CORRECTION', 'POA', 'COMMUNICATIONS',
    'BRUNEI ACKNOWLEDGEMENT', 'BRUNEI COMPLETED REPORT', 'WIPO 1ST PUB',
    'WIPO 2ND PUB', 'FOR PICK-UP CERT', 'NNP PUB', 'NOTICE OF NON-PAYMENT',
    'LPO PUBLICATION', 'FORFEITED (processed annuity)', 'REVIEWED COC',
    'VERIFIED PRU OUTPUT'
]

# Initialize an empty list to store rows for the monthly_tally_df
tally_rows = []

# This is where you'll need to integrate Flask's file handling.
# Instead of uploaded_files = files.upload(), you'll get files from Flask's request.files
# Example (within a Flask route function):
# from flask import request
# uploaded_files = request.files.getlist('file_input_name') # 'file_input_name' is the name of your file input in HTML

# Assuming you have a list of uploaded file objects (each with a filename and data stream)
# Loop through each uploaded file:
# for uploaded_file in uploaded_files:
#     filename = uploaded_file.filename
#     file_content = uploaded_file.read() # Read the file content

#     # Read the Excel file content into a pandas DataFrame
#     try:
#         df = pd.read_excel(io.BytesIO(file_content))
#     except Exception as e:
#         print(f"Error reading Excel file {filename}: {e}")
#         continue # Skip to the next file if reading fails

#     # Extract data (adjust slicing if necessary based on your file structure)
#     try:
#         names = df.iloc[2:17, 0]
#         office_status = df.iloc[2:17, 1]
#         tax_data = df.iloc[2:17, 2:47].copy()
#     except IndexError as e:
#         print(f"Error extracting data from {filename}: {e}. Check row/column ranges.")
#         continue

#     # Assign tax data columns
#     if len(tax_data.columns) == len(tax_data_columns):
#         tax_data.columns = tax_data_columns
#     else:
#         print(f"Warning: Number of tax data columns in {filename} ({len(tax_data.columns)}) does not match expected ({len(tax_data_columns)}). Skipping tax data processing for this file.")
#         tax_data = pd.DataFrame(0, index=tax_data.index, columns=tax_data_columns) # Create an empty tax data df to avoid errors

#     # Extract date from filename using regex
#     month_name = None
#     match = re.search(r'Summary\s(.*)\.xlsx', filename)
#     if match:
#         date_str = match.group(1)
#         try:
#             date_obj = datetime.strptime(date_str, '%b %d, %Y')
#             month_name = date_obj.strftime('%b').upper()
#         except ValueError:
#             print(f"Could not parse date from filename: {filename}")
#     else:
#         print(f"Could not extract date from filename: {filename}")

#     if month_name:
#         # Create structured DataFrame for the current file
#         structured_df = pd.DataFrame()
#         # Ensure names and office_status have the same length as tax_data rows after slicing
#         min_len = min(len(names), len(office_status), len(tax_data))
#         structured_df['Name'] = names.iloc[:min_len].values
#         structured_df['Office Status'] = office_status.iloc[:min_len].values

#         # Concatenate data - ensure indices align or reset index
#         structured_df = pd.concat([structured_df.reset_index(drop=True), tax_data.reset_index(drop=True)], axis=1)

#         # Add month as a new column
#         structured_df['Month'] = month_name

#         # --- Aggregation and Tally Population for the current file ---
#         # Group by 'Name' and 'Month' and sum the tax data columns
#         aggregated_df = structured_df.groupby(['Name', 'Month'])[tax_data_columns].sum().reset_index()

#         # Populate tally_rows from aggregated_df
#         for index, row in aggregated_df.iterrows():
#             name = row['Name']
#             current_month = row['Month']

#             # Find or create the "Total Processed Tasks" row for this processor
#             total_row = next((item for item in tally_rows if item['PROCESSOR'] == name and item['PROCESSES'] == 'Total Processed Tasks'), None)
#             if not total_row:
#                 total_row = {'PROCESSES': 'Total Processed Tasks', 'PROCESSOR': name,
#                              'JUL': 0, 'AUG': 0, 'SEP': 0, 'OCT': 0, 'NOV': 0, 'DEC': 0,
#                              'TOTAL PROCESSED': 0}
#                 tally_rows.append(total_row)

#             # Update total processed for the month and grand total
#             total_processed_this_month = row[tax_data_columns].sum()
#             if current_month in total_row:
#                 total_row[current_month] += total_processed_this_month
#                 total_row['TOTAL PROCESSED'] += total_processed_this_month
#             else:
#                  print(f"Warning: Month {current_month} from filename is not in the tally table columns. Skipping.")


#             # Add or update rows for each specific task
#             for task in tax_data_columns:
#                 task_row = next((item for item in tally_rows if item['PROCESSOR'] == name and item['PROCESSES'] == task), None)
#                 if not task_row:
#                     task_row = {'PROCESSES': task, 'PROCESSOR': name,
#                                 'JUL': 0, 'AUG': 0, 'SEP': 0, 'OCT': 0, 'NOV': 0, 'DEC': 0,
#                                 'TOTAL PROCESSED': 0}
#                     tally_rows.append(task_row)

#                 # Update task count for the month and grand total
#                 task_count = row[task]
#                 if current_month in task_row:
#                     task_row[current_month] += task_count
#                     task_row['TOTAL PROCESSED'] += task_count
#                 else:
#                     print(f"Warning: Month {current_month} from filename is not in the tally table columns for task {task}. Skipping.")


# After processing all files, create the final DataFrame
# monthly_tally_df = pd.DataFrame(tally_rows)

# # Filter the DataFrame to only show rows where 'TOTAL PROCESSED' is greater than 0
# filtered_monthly_tally_df = monthly_tally_df[monthly_tally_df['TOTAL PROCESSED'] > 0].reset_index(drop=True)

# # Now filtered_monthly_tally_df can be rendered as an HTML table in Flask
# # Example using pandas to_html():
# # html_table = filtered_monthly_tally_df.to_html()
