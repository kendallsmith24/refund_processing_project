'''
The following Python script extracts all data from PDF sales receipts, reorganizes and recalculates the totals to determine if refunds/credits should be issued.
It later writes the results to Microsoft Excel spreadsheets.
'''

# IMPORT LIBRARIES
from pathlib import Path
import os
import re
import pdfplumber
import numpy as np
import pandas as pd
from datetime import datetime
import xlsxwriter


#=============================================================================================================================================================================================
#=============================================================================================================================================================================================

input_folder_path = Path(INPUT_FOLDER_PATH) # Folder path where pdf receipts are held
output_folder_path = OUTPUT_FOLDER_PATH # Destination folder path of where the output files will write to 

file_name = os.path.basename(input_folder_path)
file_name = os.path.splitext(file_name)[0]

all_receipt_details = [] # List to hold all receipt details
all_receipt_comparisons = [] # List to hold all receipt comparisons
all_receipt_meta = [] # List to hold all receipt metadata 
processed_receipt_numbers = [] # List to hold processed receipt numbers for duplicate detection

input_folder_path = Path(input_folder_path)
for file_name in os.listdir(input_folder_path):
    if file_name.lower().endswith(".pdf"):
        full_path = os.path.join(input_folder_path, file_name)
            
        file_name = os.path.splitext(file_name)[0]
        try:
            with pdfplumber.open(full_path) as pdf:
                # Iterate through pages
                for page in pdf.pages:
                    # Extract tables from the current page
                    tables = page.extract_tables()
                    for table in tables:
                        # Convert the list of lists to a pandas DataFrame for easier manipulation
                        data = pd.DataFrame(table[1:], columns=table[0])
            print(f"Initiating receipt data retrieval for {file_name}...")
            pdf_layout = page.extract_text(layout=True) # Assign the true layout of the PDF to a variable (print(pdf_layout) shows the layout)
            data = data.replace(r'^\s*$', np.nan, regex=True).dropna(how='all') # Remove any rows where all columns contain either spaces or NaN values
            data = data.replace(r'\$', '', regex=True) # Remove Dollar Sign ($) from data

            # ==========================================================================================================================================
            # == DATA EXTRACTION
            # ==========================================================================================================================================
                
            # RECEIPT DETAIL CHECK DATAFRAME CREATION
            # Creates dataframe from data by dropping any rows that have NaNs in mentioned columns (removes rows that contain total discount, subtotal, sales tax and total - only leaves line items)
            receipt_detail_check = data.dropna(subset=['QTY', 'ITEM #', 'DESCRIPTION'])
            
            
            # REGEX PATTERNS
            date_pattern = r'(\d{1,2}/\d{1,2}/\d{2})' # Returns date on receipt
            receipt_pattern = r'RECEIPT #[^\n]*?\b(\d{3,4})\b' # Returns receipt number (3 to 4 digit number after keyword "RECEIPT #")
            customer_pattern = r'ID NO.[^\n]*?([A-Za-z0-9]{6})' # Returns customer id number (First instance of 6-digit alphanumeric after keyword "ID NO.")
            payment_pattern = r'(Credit|Debit|Cash)' # Returns payment type keywords
            check_no_pattern = r'(.+?)\s*(?=(Credit|Debit|Cash))' # Returns value before payment type keywords (check # exists on same row as payment method on PDF)
            job_pattern = r'(?:Credit|Debit|Cash)\s*(.+)' # Returns value after payment type keywords (job exists on same row as payment method on PDF)

            # DATA EXTRACTION BLOCKS ==================================

            # DATE
            # Assigns date extraction to column in receipt_detail_check dataframe
            date_match = re.search(date_pattern, pdf_layout)
            if date_match:
                date = date_match.group(1)
            else:
                date = np.nan
            receipt_detail_check['date_of_receipt'] = date
            receipt_detail_check['date_of_receipt'] = pd.to_datetime(receipt_detail_check['date_of_receipt'], format='%m/%d/%y')
            
            
            # RECEIPT
            # Assigns receipt number extraction to column in receipt_detail_check dataframe
            receipt_match = re.search(receipt_pattern, pdf_layout)
            if receipt_match:
                receipt = receipt_match.group(1)
                if receipt not in processed_receipt_numbers: # Check if receipt number has not already been processed
                    processed_receipt_numbers.append(receipt) # Add receipt number to list of processed receipt numbers for duplicate detection
                else: 
                    receipt = receipt + " - D" # If receipt number has already been processed, append "- D" for duplicate (prevents duplicate receipt numbers from being added to list and incorrectly flagged as duplicates)
            else:
                receipt = np.nan
            receipt_detail_check['receipt_number'] = receipt
            

            # CUSTOMER ID NUMBER
            # Assigns customer id number extraction to column in receipt_detail_check dataframe
            customer_match = re.search(customer_pattern, pdf_layout)
            if customer_match:
                customer = customer_match.group(1)
            else:
                customer = np.nan
            receipt_detail_check['customer_id_no'] = customer
            

            # CHECK NUMBER
            check_no_match = re.search(check_no_pattern, pdf_layout)
            if check_no_match:
                check = check_no_match.group(1)
            else:
                check = np.nan
            receipt_detail_check['check_no'] = check
            

            # CREATING RECEIPT META DATAFRAME
            meta_data = receipt_detail_check.loc[:, 'date_of_receipt':'check_no']


            # PAYMENT METHOD
            payment_match = re.search(payment_pattern, pdf_layout)
            if payment_match:
                payment = payment_match.group(1)
            else: 
                payment = np.nan
            meta_data['payment_method'] = payment


            # JOB
            job_match = re.search(job_pattern, pdf_layout)
            if job_match:
                job = job_match.group(1)
            else: 
                job = np.nan
            meta_data['job'] = job
            
            # Reassigning the meta_data dataframe to only the first row
            meta_data = meta_data.iloc[[0]]

            # Renaming columns
            receipt_detail_check = receipt_detail_check.rename(columns={'QTY':'qty', 'ITEM #':'item_no',
                                                                        'DESCRIPTION':'description', 'UNIT PRICE':'unit_price',
                                                                        'DISCOUNT':'discount', 'LINE TOTAL':'line_total'})

            # Reordering columns
            column_order = ['date_of_receipt', 'receipt_number', 'customer_id_no', 'check_no',
                            'qty', 'item_no', 'description', 'unit_price', 'discount', 'line_total']
            receipt_detail_check = receipt_detail_check.reindex(columns=column_order)

            # Changing data types of the quantity, unit price, discount, and line total columns
            receipt_detail_check = receipt_detail_check.astype({'qty':'int', 'unit_price':'float', 'discount':'float', 'line_total':'float'})
                        
        except Exception as e: # If error occurs continue to next file
            print(f'Error processing {file_name}: {e}. Skipping to next file...')

        #==========================================================================================================================================
        #== DATA TRANSFORMATION
        #==========================================================================================================================================    
                    
        # RECEIPT COMPARISON DATAFRAME CREATION
        receipt_comparison = pd.DataFrame(columns=['receipt_assessment_date',
                                                'date_of_receipt',
                                                'receipt_number',
                                                'subtotal_recheck',
                                                'receipt_subtotal',
                                                'discount_recheck',
                                                'receipt_discount', 
                                                'sales_tax', 
                                                'total_recheck',
                                                'receipt_total',
                                                'receipt_comparison_difference', 
                                                'receipt_totals_match',
                                                'refund_authorized', 
                                                'refund_amount'])

        receipt_comparison = receipt_comparison.assign(date_of_receipt = receipt_detail_check['date_of_receipt'].iloc[0:1], # Date of Receipt
                                                        receipt_number = receipt_detail_check['receipt_number'].iloc[0:1], # Receipt Number
                                                        subtotal_recheck = receipt_detail_check['line_total'].sum().astype(float), # Aggregating recalculated line totals in the receipt_detail_check dataframe and assigning to the receipt_comparison dataframe
                                                        discount_recheck = receipt_detail_check['discount'].sum().astype(float), # Aggregating the recalculated discounts
                                                        sales_tax = data['LINE TOTAL'].astype(float).iloc[-2], # Sales Tax Extraction
                                                        total_recheck = (receipt_detail_check['line_total'].sum().astype(float) - receipt_detail_check['discount'].sum().astype(float)) + data['LINE TOTAL'].astype(float).iloc[-2], # Recalculated Total (subtotal minus total discount plus sales tax)
                                                        receipt_subtotal = data['LINE TOTAL'].astype(float).iloc[-3], # Subtotal Extraction
                                                        receipt_discount = data['DISCOUNT'].astype(float).iloc[-4], # Total Discount Extraction
                                                        receipt_total = data['LINE TOTAL'].astype(float).iloc[-1], # Receipt Total Extraction 
                                                        receipt_comparison_difference = ((receipt_detail_check['line_total'].sum().astype(float) - receipt_detail_check['discount'].sum().astype(float)) + data['LINE TOTAL'].astype(float).iloc[-2] - data['LINE TOTAL'].astype(float).iloc[-1])) # Difference between recalculated total and receipt total

        
        # Ensuring any negative value shows a positive with parentheses
        receipt_comparison['receipt_comparison_difference'] = receipt_comparison['receipt_comparison_difference'].apply(lambda x: f'({abs(x):.2f})' if x < 0 else f'{x:.2f}')

        # Where the recalculated total and original receipt total match apply 'Yes' to receipts_match column else 'No'
        receipt_comparison['receipt_totals_match'] = np.where(receipt_comparison['total_recheck'] == receipt_comparison['receipt_total'], 'Yes', 'No')

        ############################################################################################################################################################################################################
        # REFUND/CREDIT AUTHORIZATION LOGIC
        refund_auth_conditions = [(receipt_comparison['receipt_totals_match'] == 'No') & (receipt_comparison['discount_recheck'] != receipt_comparison['receipt_discount']),
                                (receipt_comparison['receipt_totals_match'] == 'No') & (receipt_comparison['discount_recheck'] == receipt_comparison['receipt_discount']),
                                (receipt_comparison['receipt_totals_match'] == 'Yes') & (receipt_comparison['discount_recheck'] != receipt_comparison['receipt_discount']),
                                (receipt_comparison['receipt_totals_match'] == 'Yes') & (receipt_comparison['discount_recheck'] == receipt_comparison['receipt_discount'])]

        # REFUND/CREDIT AUTHORIZATION CHOICES
        refund_auth_choices = ['Yes', 'Manual Check Needed - Total Mismatch/Discount Match', 'Manual Check Needed - Total Match/Discount Mismatch', 'No']

        # Applying conditions to choices and assigning to refund_authorized column
        receipt_comparison['refund_authorized'] = np.select(refund_auth_conditions, refund_auth_choices, default="Manual Check Needed")

        ############################################################################################################################################################################################################
        # REFUND AMOUNT ASSIGNMENT LOGIC
        # Applying logic to assign refund amount based on refund authorization and whether receipt totals match or not using np.select for multiple conditions.                                                                
        refund_amt_conditions = [receipt_comparison['refund_authorized'] == 'Yes',
                                receipt_comparison['refund_authorized'].str.contains('Manual Check Needed'),
                                receipt_comparison['refund_authorized'] == 'No']

        refund_amt_choices = [receipt_comparison['receipt_comparison_difference'], 
                              receipt_comparison['receipt_comparison_difference'] + " - Pending Manual Check", 
                              receipt_comparison['receipt_comparison_difference']]

        # Applying conditions to choices and assigning to refund_amount column
        receipt_comparison['refund_amount'] = np.select(refund_amt_conditions, refund_amt_choices, default = receipt_comparison['receipt_comparison_difference'])

        ############################################################################################################################################################################################################
        # RECEIPT ASSESSMENT DATE TO RECEIPT COMPARISON DATAFRAME
        
        # Assigning value to when receipt was processed
        receipt_comparison['receipt_assessment_date'] = datetime.now().replace(microsecond=0) # .replace(microsecond=0 removes microseconds from output)

        ############################################################################################################################################################################################################
        all_receipt_details.append(receipt_detail_check) # Append receipt details to all_receipt_details list
        all_receipt_comparisons.append(receipt_comparison) # Append receipt details to all_receipt_comparisons list
        all_receipt_meta.append(meta_data) # Append receipt meta to all_receipt_meta list
        
        if all_receipt_details and all_receipt_comparisons and all_receipt_meta:
            master_receipt_details = pd.concat(all_receipt_details, ignore_index=True)
            master_receipt_comparisons = pd.concat(all_receipt_comparisons, ignore_index=True)
            master_receipt_meta = pd.concat(all_receipt_meta, ignore_index=True)
           
            # If duplicate receipt numbers are detected, assign "Manual Check Needed - Duplicate Receipt Numbers Detected" to the refund_authorized column for those receipts in the master_receipt_comparisons dataframe
            master_receipt_comparisons.loc[master_receipt_comparisons['receipt_number'].str.contains(' - D', na=False), 'refund_authorized'] = 'Manual Check Needed - Duplicate Receipt Numbers Detected'
            
            # Check for missing receipt numbers and assign "Manual Check Needed - Receipt Number Missing" to the refund_authorized column for those receipts in the master_receipt_comparisons dataframe  
            master_receipt_comparisons.loc[master_receipt_comparisons['receipt_number'].isna(), 'refund_authorized'] = 'Manual Check Needed - Receipt Number Missing'
            
            # Check for missing date of receipt values and assign "Manual Check Needed - Date of Receipt Missing" to the refund_authorized column for those receipts in the master_receipt_comparisons dataframe
            master_receipt_comparisons.loc[master_receipt_comparisons['date_of_receipt'].isna(), 'refund_authorized'] = 'Manual Check Needed - Date of Receipt Missing'
            
            # For any receipts that have been assigned "Manual Check Needed" in the refund_authorized column, assign the refund_amount column to show the receipt_comparison_difference value with " - Pending Manual Check" appended to the end of it in the master_receipt_comparisons dataframe
            master_receipt_comparisons.loc[master_receipt_comparisons['refund_authorized'].str.contains('Manual Check Needed'), 'refund_amount'] = master_receipt_comparisons['receipt_comparison_difference'].astype(str) + " - Pending Manual Check"

            with pd.ExcelWriter(f'{output_folder_path}\consolidated_receipt_data.xlsx',  
                engine="xlsxwriter", datetime_format='m/dd/yyyy', date_format='m/dd/yyyy') as writer:
                master_receipt_meta.to_excel(writer, sheet_name='receipt_metadata', index=False)
                master_receipt_details.to_excel(writer, sheet_name='consolidated_receipt_details', index=False)
                master_receipt_comparisons.query("refund_authorized.str.contains('Manual Check Needed')").to_excel(writer, engine="xlsxwriter", sheet_name='manual_checks', index=False)

                worksheet_1 = writer.sheets['receipt_metadata']
                worksheet_2 = writer.sheets['consolidated_receipt_details']
                worksheet_3 = writer.sheets['manual_checks']
                worksheet_1.autofit()
                worksheet_2.autofit()
                worksheet_3.autofit()

        ############################################################################################################################################################################################################
        # WRITING DATA TO EXCEL WORKBOOK AND FORMATTING
        # Creating and naming Workbook and Worksheet objects

        with xlsxwriter.Workbook(f'{output_folder_path}\{file_name}.xlsx', {'default_date_format':'m/dd/yyyy'}) as workbook:

            # xlsxwriter does not allow for writing of NaN values to cells and will throw an error if attempted. 
            # The following lines convert any NaN values in the date_of_receipt and receipt_number columns to None so that they will be written to the Excel file without error. 
            # This also allows for easier identification of missing values in the Excel file as the cells will be blank instead of showing "NaN". 
            meta_data['date_of_receipt'] = meta_data['date_of_receipt'].astype(object).where(meta_data['date_of_receipt'].notna(), None) # value will need to be converted back to datetime format in excel and in dataframe if further processing is needed
            receipt_detail_check['date_of_receipt'] = receipt_detail_check['date_of_receipt'].astype(object).where(receipt_detail_check['date_of_receipt'].notna(), None) # value will need to be converted back to datetime format in excel and in dataframe if further processing is needed                  
            receipt_detail_check['receipt_number'] = receipt_detail_check['receipt_number'].astype(object).where(receipt_detail_check['receipt_number'].notna(), None)
            receipt_comparison['receipt_number'] = receipt_comparison['receipt_number'].astype(object).where(receipt_comparison['receipt_number'].notna(), None)
            
            # Adding logic to assign "Manual Check Needed - Receipt Number Missing" and "Manual Check Needed - Date of Receipt Missing" to the refund_authorized column in the receipt_comparison dataframe
            # for any receipts that are missing receipt numbers or date of receipt values (NaN values that have been converted to None values for Excel writing purposes)
            receipt_comparison.loc[receipt_comparison['receipt_number'].isna(), 'refund_authorized'] = 'Manual Check Needed - Receipt Number Missing'
            receipt_comparison.loc[receipt_comparison['date_of_receipt'].isna(), 'refund_authorized'] = 'Manual Check Needed - Date of Receipt Missing'

            # Adding new worksheet to contain the receipt_detail_check information
            worksheet = workbook.add_worksheet('Receipt Assessment')  

            # Creating list of receipt_detail_check column headers
            receipt_assessment_col_headers = ['receipt_assessment_date',
                                            'receipt_number',
                                            'refund_authorized',
                                            'refund_amount']

            # Setting column headers to the receipt_detail_check worksheet
            for col_num, header in enumerate(receipt_assessment_col_headers):
                worksheet.write(0, col_num, header)

                # Writing data to columns in the receipt_detail_check worksheet (row, column, data[column])
                worksheet.write_column(1, 0, receipt_comparison['receipt_assessment_date'])
                worksheet.write_column(1, 1, receipt_comparison['receipt_number'])
                worksheet.write_column(1, 2, receipt_comparison['refund_authorized'])
                worksheet.write_column(1, 3, receipt_comparison['refund_amount']) 


            # Creating list of receipt_detail_check column headers
            receipt_meta_col_headers = ['date_of_receipt',
                                        'customer_id_no',
                                        'check_no',
                                        'payment_method',
                                        'job']

            # Setting column headers to the receipt_detail_check worksheet
            for col_num, header in enumerate(receipt_meta_col_headers):
                worksheet.write(4, col_num, header)

                # Writing data to columns in the receipt_detail_check worksheet (row, column, data[column])
                worksheet.write_column(5, 0, meta_data['date_of_receipt'])
                worksheet.write_column(5, 1, meta_data['customer_id_no'])
                worksheet.write_column(5, 2, meta_data['check_no'])
                worksheet.write_column(5, 3, meta_data['payment_method'])
                worksheet.write_column(5, 4, meta_data['job'])


            # Creating list of receipt_comparison column headers
            receipt_comparison_col_headers = ['subtotal_recheck',
                                            'receipt_subtotal',
                                            'discount_recheck',
                                            'receipt_discount', 
                                            'sales_tax', 
                                            'total_recheck',
                                            'receipt_total',
                                            'receipt_comparison_difference', 
                                            'receipt_totals_match']


            # Setting column headers to the receipt_comparison worksheet
            for col_num, header in enumerate(receipt_comparison_col_headers):
                worksheet.write(7, col_num, header)

                # Writing data to columns in the receipt_comparison worksheet (row, column, data[column])
                worksheet.write_column(8, 0, receipt_comparison['subtotal_recheck'])
                worksheet.write_column(8, 1, receipt_comparison['receipt_subtotal'])
                worksheet.write_column(8, 2, receipt_comparison['discount_recheck'])
                worksheet.write_column(8, 3, receipt_comparison['receipt_discount'])
                worksheet.write_column(8, 4, receipt_comparison['sales_tax'])
                worksheet.write_column(8, 5, receipt_comparison['total_recheck'])
                worksheet.write_column(8, 6, receipt_comparison['receipt_total'])
                worksheet.write_column(8, 7, receipt_comparison['receipt_comparison_difference'])
                worksheet.write_column(8, 8, receipt_comparison['receipt_totals_match']) 
            

            # Creating list of receipt_detail_check column headers
            receipt_detail_col_headers = ['qty',
                                        'item_no', 
                                        'description',
                                        'unit_price',
                                        'discount',
                                        'line_total']


            # Setting column headers to the receipt_detail_check worksheet
            for col_num, header in enumerate(receipt_detail_col_headers):
                worksheet.write(10, col_num, header)

                # Writing data to columns in the receipt_detail_check worksheet (row, column, data[column])
                worksheet.write_column(11, 0, receipt_detail_check['qty'])
                worksheet.write_column(11, 1, receipt_detail_check['item_no'])
                worksheet.write_column(11, 2, receipt_detail_check['description'])
                worksheet.write_column(11, 3, receipt_detail_check['unit_price'])
                worksheet.write_column(11, 4, receipt_detail_check['discount'])
                worksheet.write_column(11, 5, receipt_detail_check['line_total'])       
                    
            # Shade column headers grey 
            column_name_shading = workbook.add_format({'bg_color':'D3D3D3'})
            worksheet.write_row(0, 0, receipt_assessment_col_headers, column_name_shading)
            worksheet.write_row(4, 0, receipt_meta_col_headers, column_name_shading)
            worksheet.write_row(7, 0, receipt_comparison_col_headers, column_name_shading)
            worksheet.write_row(10, 0, receipt_detail_col_headers, column_name_shading)

            # Automatic resizing of columns (autofit) 
            worksheet.autofit()

            print("Data has been written to Excel file. Complete.")

            # Adding new worksheet to contain the receipt_detail_check information
            worksheet = workbook.add_worksheet('Receipt_Details_Only')

            # Creating list of receipt_detail_check column headers
            details_only_col_headers = ['date_of_receipt',
                                        'receipt_number',
                                        'customer_id_no',
                                        'check_no',
                                        'qty',
                                        'item_no', 
                                        'description',
                                        'unit_price',
                                        'discount',
                                        'line_total']

            # Setting column headers to the receipt_detail_check worksheet
            for col_num, header in enumerate(details_only_col_headers):
                worksheet.write(0, col_num, header)

                # Writing data to columns in the receipt_detail_check worksheet (row, column, data[column])
                worksheet.write_column(1, 0, receipt_detail_check['date_of_receipt'])
                worksheet.write_column(1, 1, receipt_detail_check['receipt_number'])
                worksheet.write_column(1, 2, receipt_detail_check['customer_id_no'])
                worksheet.write_column(1, 3, receipt_detail_check['check_no'])
                worksheet.write_column(1, 4, receipt_detail_check['qty'])
                worksheet.write_column(1, 5, receipt_detail_check['item_no'])
                worksheet.write_column(1, 6, receipt_detail_check['description'])
                worksheet.write_column(1, 7, receipt_detail_check['unit_price'])
                worksheet.write_column(1, 8, receipt_detail_check['discount'])
                worksheet.write_column(1, 9, receipt_detail_check['line_total'])       
                    
            # Shade column headers grey 
            column_name_shading = workbook.add_format({'bg_color':'D3D3D3'})
            worksheet.write_row(0, 0, details_only_col_headers, column_name_shading)

            # Automatic resizing of columns (autofit) 
            worksheet.autofit()

            # Freeze the top row 
            worksheet.freeze_panes(1, 0)

            print("All data has been written to Excel files. Complete.")
