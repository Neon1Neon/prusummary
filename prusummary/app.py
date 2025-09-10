from flask import Flask, request, render_template, redirect, url_for, flash
import pandas as pd
import re
import io
from datetime import datetime
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # Change this in production
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

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

def process_excel_files(uploaded_files):
    """Process uploaded Excel files and return the monthly tally DataFrame."""
    tally_rows = []
    
    for uploaded_file in uploaded_files:
        filename = uploaded_file.filename
        file_content = uploaded_file.read()
        
        # Read the Excel file content into a pandas DataFrame
        try:
            df = pd.read_excel(io.BytesIO(file_content))
        except Exception as e:
            flash(f"Error reading Excel file {filename}: {e}")
            continue
        
        # Extract data (adjust slicing if necessary based on your file structure)
        try:
            names = df.iloc[2:17, 0]
            office_status = df.iloc[2:17, 1]
            tax_data = df.iloc[2:17, 2:47].copy()
        except IndexError as e:
            flash(f"Error extracting data from {filename}: {e}. Check row/column ranges.")
            continue
        
        # Assign tax data columns
        if len(tax_data.columns) == len(tax_data_columns):
            tax_data.columns = tax_data_columns
        else:
            flash(f"Warning: Number of tax data columns in {filename} ({len(tax_data.columns)}) does not match expected ({len(tax_data_columns)}). Skipping tax data processing for this file.")
            tax_data = pd.DataFrame(0, index=tax_data.index, columns=tax_data_columns)
        
        # Extract date from filename using regex
        month_name = None
        match = re.search(r'Summary\s(.*)\.xlsx', filename)
        if match:
            date_str = match.group(1)
            try:
                date_obj = datetime.strptime(date_str, '%b %d, %Y')
                month_name = date_obj.strftime('%b').upper()
            except ValueError:
                flash(f"Could not parse date from filename: {filename}")
        else:
            flash(f"Could not extract date from filename: {filename}")
        
        if month_name:
            # Create structured DataFrame for the current file
            structured_df = pd.DataFrame()
            min_len = min(len(names), len(office_status), len(tax_data))
            structured_df['Name'] = names.iloc[:min_len].values
            structured_df['Office Status'] = office_status.iloc[:min_len].values
            
            # Concatenate data - ensure indices align or reset index
            structured_df = pd.concat([structured_df.reset_index(drop=True), tax_data.reset_index(drop=True)], axis=1)
            
            # Add month as a new column
            structured_df['Month'] = month_name
            
            # Group by 'Name' and 'Month' and sum the tax data columns
            aggregated_df = structured_df.groupby(['Name', 'Month'])[tax_data_columns].sum().reset_index()
            
            # Populate tally_rows from aggregated_df
            for index, row in aggregated_df.iterrows():
                name = row['Name']
                current_month = row['Month']
                
                # Find or create the "Total Processed Tasks" row for this processor
                total_row = next((item for item in tally_rows if item['PROCESSOR'] == name and item['PROCESSES'] == 'Total Processed Tasks'), None)
                if not total_row:
                    total_row = {'PROCESSES': 'Total Processed Tasks', 'PROCESSOR': name,
                                'JUL': 0, 'AUG': 0, 'SEP': 0, 'OCT': 0, 'NOV': 0, 'DEC': 0,
                                'TOTAL PROCESSED': 0}
                    tally_rows.append(total_row)
                
                # Update total processed for the month and grand total
                total_processed_this_month = row[tax_data_columns].sum()
                if current_month in total_row:
                    total_row[current_month] += int(total_processed_this_month)
                    total_row['TOTAL PROCESSED'] += int(total_processed_this_month)
                else:
                    flash(f"Warning: Month {current_month} from filename is not in the tally table columns. Skipping.")
                
                # Add or update rows for each specific task
                for task in tax_data_columns:
                    task_row = next((item for item in tally_rows if item['PROCESSOR'] == name and item['PROCESSES'] == task), None)
                    if not task_row:
                        task_row = {'PROCESSES': task, 'PROCESSOR': name,
                                   'JUL': 0, 'AUG': 0, 'SEP': 0, 'OCT': 0, 'NOV': 0, 'DEC': 0,
                                   'TOTAL PROCESSED': 0}
                        tally_rows.append(task_row)
                    
                    # Update task count for the month and grand total
                    task_count = row[task]
                    if current_month in task_row:
                        task_row[current_month] += int(task_count)
                        task_row['TOTAL PROCESSED'] += int(task_count)
                    else:
                        flash(f"Warning: Month {current_month} from filename is not in the tally table columns for task {task}. Skipping.")
    
    # Create the final DataFrame
    monthly_tally_df = pd.DataFrame(tally_rows)
    
    # Filter the DataFrame to only show rows where 'TOTAL PROCESSED' is greater than 0
    if not monthly_tally_df.empty:
        filtered_monthly_tally_df = monthly_tally_df[monthly_tally_df['TOTAL PROCESSED'] > 0].reset_index(drop=True)
        return filtered_monthly_tally_df
    else:
        return pd.DataFrame()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files' not in request.files:
        flash('No files selected')
        return redirect(url_for('index'))
    
    files = request.files.getlist('files')
    
    if not files or all(file.filename == '' for file in files):
        flash('No files selected')
        return redirect(url_for('index'))
    
    # Process the uploaded files
    result_df = process_excel_files(files)
    
    if result_df.empty:
        flash('No data could be processed from the uploaded files')
        return redirect(url_for('index'))
    
    # Convert DataFrame to HTML for display
    html_table = result_df.to_html(classes='table table-striped table-bordered', table_id='results-table')
    
    return render_template('results.html', table=html_table, row_count=len(result_df))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)