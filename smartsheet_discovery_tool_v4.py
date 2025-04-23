
import PySimpleGUI as sg
import requests

def get_base_url(region):
    return "https://api.smartsheet.eu/2.0/" if region == "Europe" else "https://api.smartsheet.com/2.0/"

def get_sheets(api_key, base_url):
    headers = {"Authorization": f"Bearer {api_key}", "Accept": "application/json"}
    response = requests.get(f"{base_url}sheets", headers=headers)
    if response.status_code == 200:
        return {sheet['name']: sheet['id'] for sheet in response.json().get('data', [])}
    else:
        sg.popup_error(f"Error fetching sheets: {response.status_code}\n{response.text}")
        return {}

def get_columns(api_key, base_url, sheet_id):
    headers = {"Authorization": f"Bearer {api_key}", "Accept": "application/json"}
    response = requests.get(f"{base_url}sheets/{sheet_id}", headers=headers)
    if response.status_code == 200:
        columns = response.json().get('columns', [])
        return {col['title']: col['id'] for col in columns}
    else:
        sg.popup_error(f"Error fetching columns: {response.status_code}\n{response.text}")
        return {}

def get_summary_fields(api_key, base_url, sheet_id):
    headers = {"Authorization": f"Bearer {api_key}", "Accept": "application/json"}
    response = requests.get(f"{base_url}sheets/{sheet_id}/summary", headers=headers)
    if response.status_code == 200:
        summary_data = response.json()
        if "fields" in summary_data:
            return {field['title']: field['id'] for field in summary_data['fields']}
        else:
            sg.popup("No summary fields found in this sheet.")
            return {}
    else:
        sg.popup_error(f"Error fetching summary fields: {response.status_code}\n{response.text}")
        return {}

layout = [
    [sg.Text('Smartsheet API Token:'), sg.InputText(password_char='*', key='API_KEY'), 
     sg.Text('Region:'), sg.Radio('Global', "REGION", default=False, key='GLOBAL'), sg.Radio('Europe', "REGION", default=True, key='EUROPE')],
    [sg.Button('Get Sheets'), sg.Text('Sheet List:'), sg.Combo([], key='SHEET_LIST', size=(40, 1))],
    [sg.Column([
        [sg.Button('Get Column IDs')],
        [sg.Text('Column IDs')],
        [sg.Multiline(size=(50, 20), key='COLUMN_OUTPUT')],
        [sg.Button('Export Column IDs')]
     ]),
     sg.Column([
        [sg.Button('Get Field IDs')],
        [sg.Text('Field IDs')],
        [sg.Multiline(size=(50, 20), key='FIELD_OUTPUT')],
        [sg.Button('Export Field IDs')]
     ])],
    [sg.Button('Exit')]
]

window = sg.Window('Smartsheet Discovery Tool v4', layout)

sheets = {}

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    elif event == 'Get Sheets':
        api_key = values['API_KEY']
        region = "Europe" if values['EUROPE'] else "Global"
        base_url = get_base_url(region)
        sheets = get_sheets(api_key, base_url)
        window['SHEET_LIST'].update(values=list(sheets.keys()))
        sg.popup(f"✅ Sheets Retrieved from {region}. Select a sheet from the dropdown.")
    elif event == 'Get Column IDs':
        sheet_name = values['SHEET_LIST']
        if sheet_name:
            region = "Europe" if values['EUROPE'] else "Global"
            base_url = get_base_url(region)
            columns = get_columns(values['API_KEY'], base_url, sheets[sheet_name])
            output = "\n".join([f"{name}: {id}" for name, id in columns.items()]) or "No Columns Found."
            window['COLUMN_OUTPUT'].update(output)
    elif event == 'Get Field IDs':
        sheet_name = values['SHEET_LIST']
        if sheet_name:
            region = "Europe" if values['EUROPE'] else "Global"
            base_url = get_base_url(region)
            fields = get_summary_fields(values['API_KEY'], base_url, sheets[sheet_name])
            output = "\n".join([f"{name}: {id}" for name, id in fields.items()]) or "No Fields Found."
            window['FIELD_OUTPUT'].update(output)
    elif event == 'Export Column IDs':
        output = values['COLUMN_OUTPUT']
        with open('column_ids_output.txt', 'w') as f:
            f.write(output)
        sg.popup('Column IDs exported to column_ids_output.txt')
    elif event == 'Export Field IDs':
        output = values['FIELD_OUTPUT']
        with open('field_ids_output.txt', 'w') as f:
            f.write(output)
        sg.popup('Field IDs exported to field_ids_output.txt')

window.close()
