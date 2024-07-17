from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, KeepTogether, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import datetime
import pandas as pd
import os

def custom_sort(df, account_order, week_order, day_order):
    df['Account'] = pd.Categorical(df['Account'], categories=account_order, ordered=True)
    df['Week'] = pd.Categorical(df['Week'], categories=week_order, ordered=True)
    df['Required Day'] = pd.Categorical(df['Required Day'], categories=day_order, ordered=True)
    df.sort_values(by=['Account', 'Week', 'Required Day'], inplace=True)
    return df

def create_header_table(logo_path, styles, page_number, account_name):
    header_style = styles['Heading2']
    header_style.alignment = 1  # Center alignment
    date_style = styles['Normal']
    date_style.alignment = 2  # Right alignment
    header_data = [
        [Image(logo_path, width=1*inch, height=0.65*inch),
         Paragraph(f'ECAM ENGINEERING LIMITED - PICK LIST - {account_name}', header_style),
         Paragraph(f'Created on: {datetime.datetime.now().strftime("%d/%m/%Y")} Page {page_number}', date_style)]
    ]
    header_table = Table(header_data, colWidths=[1*inch, None, 2.75*inch])
    header_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    return header_table

def onPage(canvas, doc, account_name, logo_path):
    header_table = create_header_table(logo_path, getSampleStyleSheet(), doc.page, account_name)
    header_table.wrap(doc.width, doc.topMargin)
    header_table.drawOn(canvas, doc.leftMargin, doc.height + doc.topMargin - 0.45*inch)

def generate_separate_picklist_pdf(excel_file, output_directory_path, logo_path, separate_sheets_output_directory_path):
    df = pd.read_excel(excel_file, sheet_name='Picklist')
    df.fillna('', inplace=True)

    account_order = ['BAM002', 'BAM003', 'BAM004', 'BAM005', 'BAM007', 'BAM008', 'BAM009', 'BAM011', 'BAM018']
    week_order = ['Week - 1', 'Week - 2', 'Week - 3', 'Week - 4', 'Week - 5', 'Week - 6', 'Week - 7']
    day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

    df = custom_sort(df, account_order, week_order, day_order)

    grouped = df.groupby('Account')

    for account, group_data in grouped:
        # Skip PDF generation if the account has no rows
        if group_data.empty:
            print(f"Skipping PDF generation for '{account}' as it has no rows.")
            continue

        elements = []
        current_datetime = datetime.datetime.now().strftime('%d_%m_%Y_%H%M')
        account_pdf_filename = f"{account}_Picklist_{current_datetime}.pdf"
        account_pdf_path = os.path.join(separate_sheets_output_directory_path, account_pdf_filename)
        account_pdf = SimpleDocTemplate(account_pdf_path, pagesize=landscape(letter), topMargin=0.75*inch, bottomMargin=0.25*inch, leftMargin=0.25*inch, rightMargin=0.25*inch)
        styles = getSampleStyleSheet()

        account_grouped = group_data.groupby(['Week', 'Required Day'])

        previous_week = None
        for (week, required_day), group in account_grouped:
            if previous_week is not None and week != previous_week:
                elements.append(PageBreak())  # Insert a page break for a new week

            data = [group.columns.to_list()] + group.values.tolist()
            table = Table(data)
            table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('BACKGROUND', (1, 1), (1, -1), colors.lightblue),
                ('BACKGROUND', (5, 1), (5, -1), colors.lightblue),
            ]))
            # Wrap the table in a KeepTogether object to prevent splitting
            keep_table_together = KeepTogether(table)

            elements.append(keep_table_together)  # Add the KeepTogether object instead of the table directly

            # Add a 0.5 inch spacer after each table, unless it's the last table
            if (week, required_day) != list(account_grouped)[-1]:
                elements.append(Spacer(1, 0.15*inch))

            previous_week = week  # Update the previous_week variable

        def onPageWrapper(canvas, doc):
            onPage(canvas, doc, account, logo_path)

        account_pdf.build(elements, onFirstPage=onPageWrapper, onLaterPages=onPageWrapper)
        print(f"PDF for 'Account {account}' created at: {account_pdf_path}")