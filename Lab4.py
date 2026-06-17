"""
COMP 593 - Lab 4: Business Process Automation
Reads sales_data.csv and generates one formatted Excel file per ORDER ID.
Usage: python lab4_solution.py <path_to_sales_data.csv>
"""

import os
import sys
import re
from datetime import date
import pandas as pd


def main():
    csv_path = get_csv_path()
    orders_dir = create_orders_dir(csv_path)
    sales_df = read_sales_data(csv_path)

    order_ids = sales_df['ORDER ID'].unique()
    for order_id in order_ids:
        save_order_to_excel(sales_df, order_id, orders_dir)

    print(f"Done — {len(order_ids)} order file(s) saved to: {orders_dir}")


def get_csv_path():
    # REQ-1/2/3: validate CLI argument
    if len(sys.argv) < 2:
        print("Error: no CSV path provided.\nUsage: python lab4_solution.py <path_to_sales_data.csv>")
        sys.exit(1)
    csv_path = sys.argv[1]
    if not os.path.isfile(csv_path):
        print(f"Error: '{csv_path}' is not a valid file.")
        sys.exit(1)
    return csv_path


def create_orders_dir(csv_path):
    # REQ-4/5: create Orders_YYYY-MM-DD sub-directory next to the CSV
    csv_dir = os.path.dirname(os.path.abspath(csv_path))
    orders_dir = os.path.join(csv_dir, f"Orders_{date.today().isoformat()}")
    os.makedirs(orders_dir, exist_ok=True)
    return orders_dir


def read_sales_data(csv_path):
    # REQ-6: load only the columns needed; utf-8-sig strips the BOM
    cols = ['ORDER ID', 'ORDER DATE', 'ITEM NUMBER', 'PRODUCT LINE',
            'PRODUCT CODE', 'ITEM QUANTITY', 'ITEM PRICE', 'STATUS', 'CUSTOMER NAME']
    return pd.read_csv(csv_path, encoding='utf-8-sig', usecols=cols)


def save_order_to_excel(sales_df, order_id, orders_dir):
    # Filter and sort by item number (REQ-7)
    order_df = sales_df[sales_df['ORDER ID'] == order_id].copy()
    order_df.sort_values('ITEM NUMBER', inplace=True)

    # REQ-8: calculate total price per line item
    order_df['TOTAL PRICE'] = order_df['ITEM QUANTITY'] * order_df['ITEM PRICE']

    # Drop ORDER ID, reorder columns for the sheet
    output_cols = ['ORDER DATE', 'ITEM NUMBER', 'PRODUCT LINE', 'PRODUCT CODE',
                   'ITEM QUANTITY', 'ITEM PRICE', 'TOTAL PRICE', 'STATUS', 'CUSTOMER NAME']
    order_df = order_df[output_cols]

    # REQ-9: append grand total row
    grand_total = order_df['TOTAL PRICE'].sum()
    grand_total_row = pd.DataFrame(
        [['', '', '', '', 'GRAND TOTAL:', '', grand_total, '', '']],
        columns=output_cols
    )
    order_df = pd.concat([order_df, grand_total_row], ignore_index=True)

    # Write to Excel
    safe_id = re.sub(r'[^\w\-]', '_', str(order_id))
    file_path = os.path.join(orders_dir, f"order_{safe_id}.xlsx")

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        order_df.to_excel(writer, index=False, sheet_name='Sheet1')
        apply_formatting(writer.book, writer.sheets['Sheet1'], order_df, grand_total)


def apply_formatting(workbook, worksheet, order_df, grand_total):
    # REQ-10: money format for ITEM PRICE (col 5) and TOTAL PRICE (col 6)
    money_fmt = workbook.add_format({'num_format': '$#,##0.00'})

    # REQ-11: column widths per lab spec
    for col_idx, width in enumerate([11, 13, 15, 15, 15, 13, 13, 10, 30]):
        worksheet.set_column(col_idx, col_idx, width)

    # Re-write price cells with money format (skips empty grand total row cells)
    for row_idx in range(1, len(order_df)):
        item_price = order_df.iloc[row_idx - 1]['ITEM PRICE']
        total_price = order_df.iloc[row_idx - 1]['TOTAL PRICE']
        if pd.notna(item_price) and item_price != '':
            worksheet.write(row_idx, 5, item_price, money_fmt)
        if pd.notna(total_price) and total_price != '':
            worksheet.write(row_idx, 6, total_price, money_fmt)

    # Grand total cell (last row, col G)
    worksheet.write(len(order_df), 6, grand_total, money_fmt)


if __name__ == '__main__':
    main()