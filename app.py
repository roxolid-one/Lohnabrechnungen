from pandas import read_excel
from django.template import Context, Engine
from django.conf import settings
from datetime import date
import os
import pdfkit
from mail import Mail
import constants

# input constants
COMPANY = constants.COMPANY

EMPLOYEE_INFO_EXCEL_PATH = constants.EMPLOYEE_INFO_EXCEL_PATH
SALARY_EXCEL_PATH = constants.SALARY_EXCEL_PATH

TEMPLATE_PATH = constants.TEMPLATE_PATH

OUTPUT_DIR = constants.OUTPUT_DIR

YEAR = constants.YEAR


def run_app():
    # variables
    this_day = date.today()
    this_day_str = f'{this_day.day}.{this_day.month}.{this_day.year}'

    months_dict = {1: 'Januar', 2: 'Februar', 3: 'März', 4: 'April', 5: 'Mai', 6: 'Juni',
                   7: 'Juli', 8: 'August', 9: 'September', 10: 'Oktober', 11: 'November', 12: 'Dezember'}

    month_name = ''

    sender_list = []
    payment_data = []

    # setup Django
    engine = Engine(dirs=['.', ])
    settings.configure()

    # setup pdfkit
    options = {"enable-local-file-access": None}

    print(f"Das aktuelle Jahr ist {YEAR} und ist manuell einzustellen (constants.py)")
    # Read Employee list
    print("Angestellten-Daten werden gelesen...")
    employees_df = read_excel(EMPLOYEE_INFO_EXCEL_PATH, sheet_name='Angaben Angestellte', header=0, index_col=0)
    # drop empty rows
    employees_df.dropna(
        axis=0,
        how='all',
        thresh=None,
        subset=None,
        inplace=True
    )
    print("Angestellten-Daten wurden gelesen.")

    # Ask for Month
    no_valid_input = True
    while no_valid_input:
        month_nr_str = input("Für welchen Monat (Nummer) möchtest du die Lohnabrechnungen senden? ")
        try:
            month_nr = int(month_nr_str)
            month_name = months_dict[month_nr]
            no_valid_input = False
        except ValueError:
            print("Ungültig, du musst eine Zahl eingeben!")

    # Read Data
    print(f"Monatsdaten von {month_name} werden gelesen...")
    data_df = read_excel(SALARY_EXCEL_PATH, sheet_name=month_name, header=0, index_col=0)
    # drop empty rows
    data_df.dropna(
        axis=0,
        how='all',
        thresh=None,
        subset=None,
        inplace=True
    )
    # replace 'nan' with 0:
    data_df.fillna(0, inplace=True)
    print(f"Monatsdaten wurden gelesen.")

    # Template and PDF Funcions
    def html_to_pdf(employee, context_dict):
        print(f"PDF für {employee} wird erstellt...")
        template = engine.get_template(TEMPLATE_PATH)
        cont = Context(context_dict)
        output_pdf_dir = f'{OUTPUT_DIR}/Lohnabrechnungen_{COMPANY}/{YEAR}/{month_name}'
        if not os.path.exists(output_pdf_dir):
            os.makedirs(output_pdf_dir)
        pdf_file_path = f'{output_pdf_dir}/{employee}_Lohnabrechnung_{month_name}.pdf'
        pdfkit.from_string(template.render(cont), pdf_file_path, options=options)
        return pdf_file_path

    def create_context(employee):
        employee_df = data_df[employee]
        employee_info = employees_df[employee]
        context_dict = {
            'Names': employee_info['Name'],
            'Adr_line_1': employee_info['Adresszeile 1'],
            'Adr_line_2': employee_info['Adresszeile 2'],
            'AHV_Nr': employee_info['AHV-Nummer'],
            'iban': employee_info['Konto-Nummer'],
            'date_today': this_day_str,
            'month_name': month_name,
            'base_salary': employee_df['Basislohn'],
            'admin_salary': employee_df['Basislohn Admin'],
            'admin_base': employee_df['Stundenansatz Admin'],
            'admin_hours': employee_df['Stunden Admin'],
            'sales_salary': employee_df['Basislohn Verkauf'],
            'sales_base': employee_df['Stundenansatz Verkauf'],
            'sales_hours': employee_df['Stunden Verkauf'],
            'holiday_perc': employee_df['Ferienentschädigung-Satz'],
            'holiday': employee_df['Ferienentschädigung'],
            'sales_total': round(employee_df['Brutto Stundelöhne'], 2),
            'other_salaries_name': list(data_df.index.values)[12],
            'other_salaries': employee_df[12],
            'children_plus': round(employee_df['Kinderzulagen'], 2),
            'other_plus_name': list(data_df.index.values)[14],
            'other_plus': round(employee_df[14], 2),
            'brutto_salary': round(employee_df['Bruttolohn'], 2),
            'calc_base': round(employee_df['Berechnungsgrundlage'], 2),
            'ahv_perc': employee_df['AHV-Satz']*100,
            'ahv_minus': round(employee_df['AHV/IV/EO-Beitrag'], 2),
            'alv_perc': round(employee_df['ALV-Satz']*100, 1),
            'alv_minus': round(employee_df['ALV-Beitrag'], 2),
            'uvg_perc': round(employee_df['UVG-Satz']*100, 3),
            'uvg_minus': round(employee_df['UVG (NBU)-Beitrag'], 2),
            'bvg_minus': round(employee_df['BVG-Beitrag'], 2),
            'ktg_perc': round(employee_df['KTG-Satz']*100, 4),
            'ktg_minus': round(employee_df['KTG-Beitrag'], 2),
            'other_minus_1_name': list(data_df.index.values)[26],
            'other_minus_1': round(employee_df[26], 2),
            'other_minus_2_name': list(data_df.index.values)[27],
            'other_minus_2': round(employee_df[27], 2),
            'corr_name': list(data_df.index.values)[28],
            'corr': round(employee_df[28], 2),
            'net_salary': round(employee_df['Nettolohn'], 2),
            'msg1': employee_df['Mitteilungen 1'],
            'msg2': employee_df['Mitteilungen 2'],
        }
        if employee_info[1] == "weiblich":
            context_dict['gender'] = 'Frauen'
        else:
            context_dict['gender'] = 'Männer'

        return context_dict

    # Create PDFs
    for employee_name in employees_df:
        context = create_context(employee_name)
        pdf_path = html_to_pdf(employee_name, context)
        sender_list.append((employee_name, employees_df[employee_name]['mail-Adresse'],
                            employees_df[employee_name]['Geschlecht'], pdf_path))
        payment_data.append((employee_name, employees_df[employee_name]['Name'],
                            employees_df[employee_name]['Konto-Nummer'], context['net_salary']))

    # Send PDFs
    print(":)")
    print("Alle Lohnabrechnungen wurden erstellt!")
    print(":)")
    print("Hier die Zahlungsdaten:")
    for payment in payment_data:
        print(f'{payment[0]}:')
        print(f'Name: {payment[1]}, IBAN: {payment[2]}, Auszahlungsbetrag: {payment[3]} CHF')
    proceed = input("Möchtest du die PDFs nun senden? (j/n): ")

    if proceed == "j":
        for recipient in sender_list:
            (name, mail_adress, gender, file_path) = recipient
            print(f"Sende Mail an {name}...")
            subject = f'{COMPANY}-Lohnabrechnung vom {month_name}'

            # mail body
            if gender == 'weiblich':
                opening = 'Liebe'
            else:
                opening = 'Lieber'
            with open('mail_body.txt', 'r') as mail_body:
                body = mail_body.read()
            body = body.replace("<courtesy>", opening)
            body = body.replace("<name>", name)
            body = body.replace("<month>", month_name)
            body = body.replace("<year>", str(YEAR))

            mail = Mail(mail_adress, subject, body, file_path)
            mail.send_mail()

        print("Alle Mails wurden gesendet!")

    # TODO: Create csv with Payment-Data

    # TODO: Progress and Raise Exeptions

    # TODO: (Gui)
