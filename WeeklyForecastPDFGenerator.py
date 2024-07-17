from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.colors import yellow, lightgreen
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, KeepTogether
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import datetime
import pandas as pd
import numpy as np
import os
from decimal import Decimal, ROUND_HALF_UP

def round_decimal(x):
    if pd.isnull(x):
        return x  # Keep NaN as is
    if isinstance(x, (float, int)):
        return Decimal(x).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP)
    return x

# Define styles in a scope accessible to both create_header and on_every_page
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='TimesNewRoman', fontName='Times New Roman', fontSize=12, leading=14))
header_style = styles['TimesNewRoman']
header_style.alignment = 1
date_style = styles['TimesNewRoman']
date_style.alignment = 2

def round_decimal(x):
    if pd.isnull(x):
        return x  # Keep NaN as is
    if isinstance(x, (float, int)):
        return Decimal(x).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP)
    return x

# Define styles in a scope accessible to both create_header and on_every_page
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='TimesRoman', fontName='Times-Roman', fontSize=12, leading=14))
header_style = styles['TimesRoman']
header_style = styles['Heading2']
header_style.alignment = 1
date_style = styles['TimesRoman']
date_style.alignment = 2

def create_header(logo_path):
    header_text = 'ECAM ENGINEERING LIMITED - WEEKLY FORECAST'
    created_on_text = 'Created on: ' + datetime.datetime.today().strftime('%d/%m/%Y')
    header_data = [
        [Image(logo_path, width=1*inch, height=0.65*inch), 
         Paragraph(header_text, style=header_style), 
         Paragraph(created_on_text, style=date_style)]
    ]
    header_table = Table(header_data, colWidths=[1*inch, None, 2*inch])
    header_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    return header_table

def on_every_page(canvas, doc, logo_path):
    canvas.saveState()
    header = create_header(logo_path)
    w, h = header.wrap(doc.width, doc.topMargin)
    header.drawOn(canvas, doc.leftMargin, doc.height + doc.topMargin - h)
    canvas.restoreState()

def generate_weekly_forecast_pdf(excel_file, output_directory_path, logo_path, sheet_name='Weekly Forecast'):
    # Read the Excel file and specified sheet into a DataFrame
    df = pd.read_excel(excel_file, sheet_name=sheet_name)

    # Calculate the grand total for each account across the specified columns (e.g., 'B:AZ')
    df['Grand Total'] = df.iloc[:, 1:].sum(axis=1)  # Sum across columns for each row

    # Initialize PDF document
    current_datetime = datetime.datetime.now().strftime('%d_%m_%Y_%H%M')
    pdf_file = os.path.join(output_directory_path, f'Weekly_Forecast_{current_datetime}.pdf')

    pdf = SimpleDocTemplate(pdf_file, pagesize=landscape(letter), topMargin=1.5*inch, bottomMargin=0.5*inch, leftMargin=0.5*inch, rightMargin=0.5*inch)

    elements = [Spacer(1, 0.25*inch)]

    # Process data and create tables
    num_weeks = len(df.columns) - 2  # Adjusted to exclude 'Account' and 'Grand Total' columns
    total_chunks = (num_weeks - 1) // 9 + 1  # Calculate total number of chunks
    chunk_count = 0  # Initialize chunk counter

    for start_week in range(1, num_weeks + 1, 9):  # Adjusted loop to correctly process chunks
        chunk_count += 1
        end_week = min(start_week + 8, num_weeks)  # Ensure end_week does not exceed num_weeks
        week_columns = ['Account'] + df.columns[start_week:end_week+1].tolist()
        
        if chunk_count == total_chunks:
            # For the last chunk, include the 'Grand Total' column
            week_columns.append('Grand Total')
        
        data_chunk = df[week_columns].copy()

        # Apply rounding with decimal precision to all numerical values
        for col in data_chunk.select_dtypes(include=[np.number]).columns:
            data_chunk[col] = data_chunk[col].apply(round_decimal)

        # Ensure only 8 or 9 rows of data are included before adding the 'Total' row
        search_services = 'BAM003'
        if search_services in data_chunk['Account'].values:
            data_chunk = data_chunk.head(9)  # Change from 8 to 9
        else:
            data_chunk = data_chunk.head(8)  # Change from 7 to 8

        # Calculate the sum for each row and add it as a 'Total' column
        if chunk_count == total_chunks:
            # For the last chunk, calculate the sum across all week columns present in the chunk, excluding the 'Grand Total' column
            data_chunk['Total'] = data_chunk.iloc[:, 1:-1].sum(axis=1).apply(round_decimal)  # Exclude 'Account' and 'Grand Total' for sum
        else:
            # For other chunks, calculate the sum across all columns after 'Account'
            data_chunk['Total'] = data_chunk.iloc[:, 1:].sum(axis=1).apply(round_decimal)

        # Calculate the sum for each column in the chunk and prepare a 'Total' row
        total_row_values = [round_decimal(data_chunk[col].sum()) for col in data_chunk.columns[1:]]  # Exclude 'Account' for sum
        total_row = pd.DataFrame([['Total'] + total_row_values], columns=data_chunk.columns)

        # Concatenate the 'Total' row
        data_chunk = pd.concat([data_chunk, total_row], ignore_index=True)


        # Prepare data for table
        data_for_table = [data_chunk.columns.tolist()] + data_chunk.values.tolist()

        # Create table for each chunk and ensure it's not split across pages
        table = Table(data_for_table)
        
        # Initialize the list of style commands
        style_commands = [
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.darkblue),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BACKGROUND', (0, -1), (-1, -1), yellow),
        ]

        # Determine the index for the 'Total' column
        total_col_index = -2 if chunk_count == total_chunks else -1
        
        # Highlight the 'Total' column in yellow
        style_commands.append(('BACKGROUND', (total_col_index, 0), (total_col_index, -1), colors.yellow))
        
        # Conditionally add the 'Grand Total' column styling for the last table
        if chunk_count == total_chunks:
            grand_total_col_index = -1  # 'Grand Total' is the last column in the last table
            # Highlight the 'Grand Total' column in green for the last table
            style_commands.append(('BACKGROUND', (grand_total_col_index, 0), (grand_total_col_index, -1), lightgreen))
        
        # Apply the styles to the table
        table.setStyle(TableStyle(style_commands))
        
        elements.append(KeepTogether(table))

        # Add a 0.5-inch space after each table
        elements.append(Spacer(1, 0.5*inch))

    # Build the PDF with the header on each page
    pdf.build(elements, onFirstPage=lambda canvas, doc: on_every_page(canvas, doc, logo_path), onLaterPages=lambda canvas, doc: on_every_page(canvas, doc, logo_path))

    print(f"PDF created for sheet: {sheet_name}")