"""--------------------ι𝐍Ⓙย𝐬𝓣ᶤςⒺ Ⓐ𝐍ＹωᕼⒺг𝐄 ᶤ𝐬 ᵃ tｈяᗴＡт ⓉＯ 𝐣υ𝔰ｔ𝐢ᶜⓔ 𝐄V乇яｙ山卄εŘ乇 --------------------
Description:
 Extracts sales data from a CSV file and exports relevant data to a formatted Excel spreadsheet.

Usage:
 python lab3.py

Parameters:
 sales_csv = Path of the CSV file.
---------------------ι𝐍Ⓙย𝐬𝓣ᶤςⒺ Ⓐ𝐍ＹωᕼⒺг𝐄 ᶤ𝐬 ᵃ tｈяᗴＡт ⓉＯ 𝐣υ𝔰ｔ𝐢ᶜⓔ 𝐄V乇яｙ山卄εŘ乇--------------------"""

from sys import argv, exit
import pandas as pd
from datetime import date
import os
import re

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

# Get path of sales data CSV file from the command line
def get_sales_csv():  
    # Check whether command line parameter provided
    num_params = len(argv) -1
    if num_params >= 1:
        sales_csv = argv[1]
        # Check whether provided parameter is valid path of file
        if os.path.isfile(sales_csv):
            return sales_csv
        else:
            print('Error: Invalid path to sales data CSV file')
            exit()
    else:
        print('Error: Missing path to sales data CSV file')
        exit()

# Create the directory to hold the individual order Excel sheets.
def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file resides.
    sales_dir = os.path.dirname(os.path.abspath(sales_csv))
    # Determine the path of the directory to hold the order data files.
    todays_date = date.today().isoformat()
    orders_dir = os.path.join(sales_dir, f'Orders_{todays_date}')
    # Create the order directory if it does not already exist.
    if not os.path.isdir(orders_dir):
        os.makedirs(orders_dir)
        return orders_dir

# Split the sales data into individual orders and save to Excel sheets.
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame.
    sales_df = pd.read_csv(sales_csv)
    # Insert a new "TOTAL PRICE" column into the DataFrame.
    sales_df.insert(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY'] * sales_df['ITEM PRICE'])

    # Remove columns from the DataFrame that are not needed.
    sales_df.drop(columns=['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace=True)
    # Group sales data by order ID and save to Excel sheets.  
    for order_id, order_df in sales_df.groupby('ORDER ID'): 
        # Remove the "ORDER ID" Column.  
        order_df.drop(columns=['ORDER ID'], inplace=True)
        # Sort the order by item number.
        order_df.sort_values(by='ITEM NUMBER', inplace=True)
        # Add the grand total row.
        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_df = pd.DataFrame({'ITEM PRICE': ['GRAND TOTAL'], 'TOTAL PRICE': [grand_total]})
        order_df = pd.concat([order_df, grand_total_df])
        export_order_to_excel(order_id, order_df, orders_dir)
        
# Export the order to Excel Spreadsheet.        
def export_order_to_excel(order_id, order_df, orders_dir):
    # Determine the file name and path for the order Excel sheet.
    customer_name = order_df['CUSTOMER NAME'].values[0]
    customer_name = re.sub(r'\W', '', customer_name)
    order_file = f'Order{order_id}_{customer_name}.xlsx'
    order_path = os.path.join(orders_dir, order_file)
    sheet_name = f'Order #{order_id}'
    # Format the Excel Spreadsheet.
    writer = pd.ExcelWriter(order_path, engine='xlsxwriter')
    order_df.to_excel(writer, index=False, sheet_name=sheet_name)
    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]
    format1 = workbook.add_format({'num_format': '$#,##0.00'})
    # Set Column Widths.
    worksheet.set_column('A:A', 11) 
    worksheet.set_column('B:B', 13)
    worksheet.set_column('C:E', 15)
    worksheet.set_column('F:G', 13, format1)
    worksheet.set_column('H:H', 10)
    worksheet.set_column('I:I', 30)
    writer.close()

if __name__ == '__main__':
    main()