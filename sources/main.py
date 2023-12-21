import PySimpleGUI as sg
import pandas as pd
import openpyxl
import re
import hashlib


db = "C:\ext_tool_DB\external_db.xlsx"
df = pd.read_excel(db)

def validate_number(number):
    if len(number) != 6 or number[-1] != '0' or not number.isdigit():
        return False
    return True
def open_input_window():
    layout_input = [
        [sg.Text('Name of External shit:'), sg.InputText(key='name')],
        [sg.Text('SGF of External shit:'), sg.InputText(key='sgf')],
        [sg.Text('Comment of External shit:'), sg.InputText(key='comment')],
        [sg.Submit(), sg.Cancel()]
    ]
    window_input = sg.Window('Input Window', layout_input, auto_size_text=(bool))
    event, values = window_input.read()
    window_input.close()

    if event == 'Submit':
        db = "C:\ext_tool_DB\external_db.xlsx"
        wb = openpyxl.load_workbook(db)
        sheet = wb.active
        if validate_number(values['sgf']):
            last_row_number = sheet.max_row
            new_row_number = last_row_number + 1
            row = [new_row_number, values['name'], values['sgf'], values['comment']]
            sheet.append(row)
            wb.save(db)
        else:
            sg.Popup('Invalid input!', 'The SGF value should be a 6-digit number ending with 0.')
            open_input_window()
    window_input.close()

def open_view_window():
    layout_view = [
        [sg.Text('Name of External shit:'), sg.InputText()
         ],
        [sg.Text('SGF of External shit:'), sg.InputText()
         ],
        [sg.Text('Comment of External shit:'), sg.InputText()
         ],
        [sg.Submit(), sg.Cancel()]
    ]
    window_view = sg.Window('View Window', layout_view, auto_size_text=(bool))
    event, values = layout_view.read()
    window_view.close()

def open_del_window():
    layout_del = [
        [sg.Text('Name of External shit:'), sg.InputText()
         ],
        [sg.Text('SGF of External shit:'), sg.InputText()
         ],
        [sg.Text('Comment of External shit:'), sg.InputText()
         ],
        [sg.Submit(), sg.Cancel()]
    ]
    window_del = sg.Window('Del Window', layout_del, auto_size_text=(bool))
    event, values = layout_del.read()
    window_del.close()

layout_enter = [
    [sg.Button('Input window', size=(40, 10))
     ],
    [sg.Button('View window', size=(40, 10))
     ],
    [sg.Button('Del window', size=(40, 10))
     ],
    [sg.Button('Exit', size=(40, 10))
     ],
]
window = sg.Window('External Tool', layout_enter, size=(300, 600), )

while True:  # The Event Loop
    event, values = window.read()
    # print(event, values) #debug
    if event in (None, 'Exit', 'Cancel'):
        break
    elif event == 'Input window':
        open_input_window()
    elif event == 'View window':
        open_view_window()
    elif event == 'Del window':
        open_del_window()
    elif event == 'Exit':
        window.close()
window.close()
