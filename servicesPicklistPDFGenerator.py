from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, KeepTogether, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import datetime
import pandas as pd
import os
from configparser import ConfigParser

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
         Paragraph(f'ECAM ENGINEERING LIMITED - SERVICES PICK LIST - {account_name}', header_style),
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

def generate_services_picklist_pdf(excel_file, output_directory_path, logo_path, account_name):
    df = pd.read_excel(excel_file, sheet_name='Services Picklist')
    df.fillna('', inplace=True)
    df = df[df['Account'] == account_name]

    # Exclude rows where 'Account' column has the value 'Forecast'
    # df = df[df['Message'] != 'FORECAST']

    account_order = ['BAM003']
    week_order = ['Arrears', 'Week - 1', 'Week - 2', 'Week - 3', 'Week - 4', 'Week - 5', 'Week - 6', 'Week - 7', 'Week - 8', 'Week - 9']
    day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

    df = custom_sort(df, account_order, week_order, day_order)

    current_datetime = datetime.datetime.now().strftime('%d_%m_%Y_%H%M')
    account_pdf_filename = f"{account_name}_Services_Picklist_{current_datetime}.pdf"
    account_pdf_path = os.path.join(output_directory_path, account_pdf_filename)
    account_pdf = SimpleDocTemplate(account_pdf_path, pagesize=landscape(letter), topMargin=0.75*inch, bottomMargin=0.25*inch, leftMargin=0.25*inch, rightMargin=0.25*inch)
    styles = getSampleStyleSheet()

    elements = []

    arrears_df = df[df['Week'] == 'Arrears']
    if not arrears_df.empty:
        arrears_table = create_table(arrears_df, styles, "Outstanding Orders (Arrears)")
        elements.append(arrears_table)
        elements.append(Spacer(1, 0.15*inch))

    current_schedules_df = df[(df['Week'] == 'Week - 1') | (df['Week'] == 'Week - 2') | (df['Week'] == 'Week - 3') | (df['Week'] == 'Week - 4') | (df['Week'] == 'Week - 5') | (df['Week'] == 'Week - 6')| (df['Week'] == 'Week - 7')]
    if not current_schedules_df.empty:
        current_schedules_table = create_table(current_schedules_df, styles, "Current Schedules (4 - Weeks Range)")
        elements.append(current_schedules_table)
        elements.append(Spacer(1, 0.15*inch))

    # post_schedules_df = df[(df['Week'] == 'Week - 5') | (df['Week'] == 'Week - 6') | (df['Week'] == 'Week - 7') | (df['Week'] == 'Week - 8') | (df['Week'] == 'Week - 9')]
    # if not post_schedules_df.empty:
    #     post_schedules_table = create_table(post_schedules_df, styles, "Post 4 - Week Schedules")
    #     elements.append(post_schedules_table)

    def onPageWrapper(canvas, doc):
        onPage(canvas, doc, account_name, logo_path)

    account_pdf.build(elements, onFirstPage=onPageWrapper, onLaterPages=onPageWrapper)
    print(f"PDF for 'Account {account_name}' created at: {account_pdf_path}")

def create_table(df, styles, table_heading):
    data = [df.columns.to_list()] + df.values.tolist()
    
    heading_style = ParagraphStyle(name='table_heading', parent=styles['Normal'], alignment=1, spaceAfter=6)
    heading = Paragraph(table_heading, heading_style)
    
    table_data = [[heading]] + data
    table = Table(table_data)
    table.setStyle(TableStyle([
        ('GRID', (0, 1), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
        ('BACKGROUND', (0, 1), (-1, 1), colors.grey),
        ('TEXTCOLOR', (0, 1), (-1, 1), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (1, 2), (1, -1), colors.lightblue),
        ('BACKGROUND', (5, 2), (5, -1), colors.lightblue),
        ('SPAN', (0, 0), (-1, 0)),  # Span the heading across all columns
    ]))
    return KeepTogether(table)


def get_user_date():
    user_input = input("Enter a date in 'yyyy, m, d' format or type 'p' to use the current date: ")
    if user_input.lower() == 'p':
        return datetime.datetime.now()
    else:
        try:
            year, month, day = map(int, user_input.split(', '))
            return datetime.datetime(year, month, day)
        except ValueError:
            print("Sorry, that is in the incorrect format. Try again!")
            return get_user_date()

if __name__ == "__main__":

# Get the date from the user or use the current date
    current_date = get_user_date()
    print("The current date is:", current_date)
    # Format the current date as dd_mm_YYYY
    formatted_current_date = current_date.strftime('%d_%m_%Y')
    account_name = 'BAM003'
    config = ConfigParser()
    # Read the config.ini file
    config.read('config.ini')
    logo_path = config['PATHS']['logo_path']
    excel_file_path = config['PATHS']['excel_file_path'].replace('{current_date}', formatted_current_date)
    services_picklist_output_directory_path = config['PATHS']['services_picklist_output_directory_path']
    generate_services_picklist_pdf(excel_file_path, services_picklist_output_directory_path, logo_path, account_name)

