from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.colors import yellow
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import datetime
import pandas as pd
import numpy as np
import os

def generate_monthly_forecast_pdf(excel_file, output_directory_path, logo_path, sheet_name='Monthly Forecast'):
    # Read the Excel file and specified sheet into a DataFrame
    df = pd.read_excel(excel_file, sheet_name=sheet_name)

    # Replace NaN values with an empty string
    df.fillna('', inplace=True)

    # Check if 'BAM003' exists in the 'Account' column
    has_bam003 = 'BAM003' in df['Account'].values

    # Determine the number of rows to include before adding the 'Total' row
    num_rows = 10 if has_bam003 else 9  # Change from 10 to 11 and 9 to 10


    # Ensure DataFrame has at least the specified number of rows
    if len(df) < num_rows:
        # Append empty rows until the DataFrame has the specified number of rows
        additional_rows = num_rows - len(df)
        df = pd.concat([df, pd.DataFrame([[''] * len(df.columns)] * additional_rows, columns=df.columns)], ignore_index=True)

    # Initialize 'total_values' with 'Total' and NaNs to match the DataFrame's width
    total_values = ['Total'] + [np.nan] * (len(df.columns) - 1)

    # Calculate the sum of each column, ignoring non-numeric columns and NaN values
    if df.replace('', np.nan).isnull().any().any():
        total_row = df.replace('', np.nan).select_dtypes(include=[np.number]).sum().round(2)
        for col in total_row.index:
            # Correctly assign rounded sum values to 'total_values' based on column location
            total_values[df.columns.get_loc(col)] = total_row[col]

    # Replace or insert the 'Total' row at the specified index (Row 9 or 10)
    df.iloc[num_rows - 1] = total_values


    # Calculate the sum for each row and update the last column with these sums
    # Assuming the sum should include all columns except the first (index 0)
    # df.iloc[:, -1] = df.iloc[:, 1:].replace('', 0).astype(float).sum(axis=1).round(2)
    # # After calculating the sum for each row, round the values to two decimal places
    df.iloc[:, -1] = df.iloc[:, 1:].replace('', 0).astype(float).sum(axis=1).apply(lambda x: round(x, 2))

    data_first_table = [df.columns.to_list()[:8]] + df.iloc[:, :8].applymap(lambda x: round(x, 2) if isinstance(x, (float, int)) else x).values.tolist()
    data_second_table = [df.columns.to_list()[0:1] + df.columns.to_list()[8:]] + df.iloc[:, [0] + list(range(8, df.shape[1]))].applymap(lambda x: round(x, 2) if isinstance(x, (float, int)) else x).values.tolist()


    # Create a PDF file in landscape orientation with reduced margins
    current_datetime = datetime.datetime.now().strftime('%d_%m_%Y_%H%M')
    pdf_file = os.path.join(output_directory_path, f'Monthly_Forecast_{current_datetime}.pdf')
    pdf = SimpleDocTemplate(pdf_file, pagesize=landscape(letter), topMargin=0.25*inch, bottomMargin=0.25*inch, leftMargin=0.25*inch, rightMargin=0.25*inch)

    # Get the styles and create specific styles for the heading and generated date
    styles = getSampleStyleSheet()
    header_style = styles['Heading2']  # Use a smaller heading style to reduce header size
    header_style.alignment = 1  # Center alignment
    date_style = styles['Normal']
    date_style.alignment = 2  # Right alignment

    # Create header table data
    header_data = [
        [Image(logo_path, width=1*inch, height=0.65*inch), 
         Paragraph('ECAM ENGINEERING LIMITED - MONTHLY FORECAST', header_style), 
         Paragraph('Created on: ' + datetime.datetime.today().strftime('%d/%m/%Y'), date_style)]
    ]

    # Create a table for the header
    header_table = Table(header_data, colWidths=[1*inch, None, 2*inch])

    # Ensure that the header table style does not have any grid or box lines
    header_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))

    # Calculate the width of the page and set column widths to fit the page
    page_width = letter[0]  # Width of the page in points (1 point = 1/72 inch)
    margin = 0.1 * inch  # Assuming a margin of 0.5 inch on each side
    usable_width = page_width - 0.1 * margin  # Subtract margins from both sides

    # Set column widths for the first table, dividing the usable width by the number of columns
    col_widths_first_table = [usable_width / 8] * 8

    # Set column widths for the second table, with the first column width and the remaining widths
    col_widths_second_table = [usable_width / (df.shape[1] - 7)] * (df.shape[1] - 7)
    col_widths_second_table.insert(0, col_widths_first_table[0])  # Set the width of the first column

    # Create a table for the first part of the data and add style
    first_table = Table(data_first_table, colWidths=col_widths_first_table)
    first_table.setStyle(TableStyle([
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),  # Change to desired header background color
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.darkblue),  # Change to desired header text color
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('BACKGROUND', (0, -1), (-1, -1), yellow),  # Highlight the last row with yellow
    # Add more color styles as needed
    ]))

     # Create a table for the second part of the data and add style
    second_table = Table(data_second_table, colWidths=col_widths_second_table)
    second_table.setStyle(TableStyle([
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),  # Change to desired header background color
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.darkblue),  # Change to desired header text color
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('BACKGROUND', (0, -1), (-1, -1), yellow),  # Highlight the last row with yellow
    ('BACKGROUND', (-1, 0), (-1, -1), yellow),  # Highlight the last column with yellow
    # Add more color styles as needed
    ]))

    # Build the PDF with both tables
    elements = [header_table, Spacer(1, 0.25*inch), first_table, Spacer(1, 0.25*inch), second_table]
    pdf.build(elements)
    print(f"PDF created for sheet: {sheet_name}")