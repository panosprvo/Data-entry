import PySimpleGUI as sg
import pandas as pd

sg.theme('LightBlue')

EXCEL_FILE = 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Client', size=(15, 1)), sg.Combo(['BMS', 'Ipsen', 'Organon', 'Takeda', 'Aspen', 'Gilead'], key='Client')],
    [sg.Text('Title', size=(15, 1)), sg.Combo(['Mr', 'Mrs', 'Ms', 'Miss', 'Dr', 'Other'], key='Title')],
    [sg.Text('First Name', size=(15, 1)), sg.InputText(key='First Name')],
    [sg.Text('Last Name', size=(15, 1)), sg.InputText(key='Last Name')],
    [sg.Text('Role', size=(15, 1)), sg.Combo(['Doctor', 'Pharmacist', 'MOP', 'MIRF'], key='Role')],
    [sg.Text('Address', size=(15, 1)), sg.InputText(key='Address')],
    [sg.Text('Postcode', size=(15, 1)), sg.InputText(key='Postcode')],
    [sg.Text('Phone', size=(15, 1)), sg.InputText(key='Phone')],
    [sg.Text('Fax', size=(15, 1)), sg.InputText(key='Fax')],
    [sg.Text('Email', size=(15, 1)), sg.InputText(key='Email')],
    [sg.Text('Enquiry', size=(15, 1)),
     sg.Checkbox('AE', key='AE'),
     sg.Checkbox('PQC', key='PQC'),
     sg.Checkbox('MI', key='MI')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Simple data entry form', layout)


def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved')
        clear_input()
window.close()
