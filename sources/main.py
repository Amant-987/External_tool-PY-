import PySimpleGUI as sg
import pandas as pd
import numpy as np
from datetime import datetime
import openpyxl
import self

class ExternalToolApp:
    def __init__(self, database_path):
        self.database_path = database_path
        self.df = pd.read_excel(database_path)
        self.headings = list(self.df.columns)

        sg.theme('LightBlue6')

        # Set the background color for the entire window
        window_background_color = '#74992e'  # Example color, you can change this
        # Initialize window_del as an instance variable
        self.window_del = None


        button_matrix = [
            [sg.Button('Input', key='input_btn', size=(15, 4), font=('Arial', 14)),
             sg.Button('View', key='view_btn', size=(15, 4), font=('Arial', 14))],
            [sg.Button('Delete', key='del_btn', size=(15, 4), font=('Arial', 14)),
             sg.Button('Exit', key='exit_btn', size=(15, 4), font=('Arial', 14))]
        ]

        # Text element to display SGF column content
        sgf_text_element = sg.Text("", key="sgf_text", size=(40, 10), font=('Arial', 16))

        self.layout_enter = [
            [sg.Text('External Tool', size=(30, 1), font=('Arial', 20), justification='center')],
            [sg.Text('')],  # Empty row above sg.VCenter
            button_matrix,
            [sg.Text('')],  # Empty row below sg.VCenter
        ]

        self.window = sg.Window('External Tool', self.layout_enter, size=(400, 400), resizable=True, finalize=True,
                                background_color=window_background_color)
        self.window.set_min_size((400, 400))

    def open_input_window(self):
        layout_input = [
            [sg.Text('Name of External shit:'), sg.InputText(key='Name')],
            [sg.Text('SGF of External shit:'), sg.InputText(key='SGF')],
            [sg.Text('Comment of External shit:'), sg.InputText(key='Comment')],
            [sg.Button('OK'), sg.Button('Cancel')]
        ]

        window_input = sg.Window('Input Window', layout_input)

        while True:
            event, values = window_input.read()

            if event == 'OK':
                if self.validate_input(values):
                    self.add_record(values)
                    sg.Popup('Record added successfully!')
                    break
                else:
                    sg.Popup('Invalid input! Please check the input values.')
            elif event == 'Cancel' or event == sg.WINDOW_CLOSED:
                break

        window_input.close()

    def validate_input(self, values):
        name = values['Name']
        sgf = values['SGF']
        comment = values['Comment']

        last_two_digits = sgf[-2:]  # Extract the last two digits of SGF
        first_two_digits = int(sgf[:2]) if sgf[:2].isdigit() else -1  # Extract the first two digits of SGF

        return (
                1 <= len(name) <= 20 and
                sgf.isdigit() and
                len(sgf) == 6 and
                1 <= first_two_digits <= 15 and
                last_two_digits in {'20', '30', '40', '50', '60'} and
                1 <= len(comment) <= 40
        )

    def add_record(self, values):
        # Get the current date and time
        current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Assuming 'Serial Number' is the first column in your DataFrame
        serial_number = len(self.df) + 1  # Incremented by 1 for the new row
        row = [serial_number, values['Name'], values['SGF'], values['Comment'], current_datetime]

        new_row = dict(zip(self.headings, row))
        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.df.to_excel(self.database_path, index=False)
        # Update the view window table after adding a record
        self.update_view_table()

    def open_view_window(self):
        layout_view = [
            [sg.Text('Name of External shit:'), sg.InputText(key='Name')],
            [sg.Text('SGF of External shit:'), sg.InputText(key='SGF')],
            [sg.Text('Comment of External shit:'), sg.InputText(key='Comment')],
            [sg.Button('Find', key='find_btn')],
            [sg.Table(
                values=self.df.tail(20).values.tolist(),
                headings=self.headings,
                auto_size_columns=True,
                justification='right',
                key='table'
                )]
            ]

        window_view = sg.Window('View Window', layout_view)

        while True:
                event, values = window_view.read()

                if event == sg.WIN_CLOSED:
                    break
                elif event == 'find_btn':
                    filtered_df = self.filter_data(values)
                    window_view['table'].update(values=filtered_df.values.tolist(), auto_size_columns=True)
                    self.update_sgf_text()  # Update SGF text after filtering

                window_view.close()

    def filter_data(self, values):
        name = values['Name']
        comment = values['Comment']

        # Exclude "deleted" records
        query = "pd.isna(self.df['deletion_datetime'])"

        if name:
            query += f" & self.df['Name'].str.contains('{name}')"

        if comment:
            query += f" & self.df['Comment'].str.contains('{comment}')"

        if query:
            filtered_df = self.df.query(query)
        else:
            filtered_df = self.df

        # Exclude rows with valid deletion_datetime
        filtered_df = filtered_df[pd.isna(filtered_df['deletion_datetime'])]

        return filtered_df.tail(20)

    def update_view_table(self):
        if hasattr(self, 'window_view') and self.window_view:
            filtered_df = self.df.tail(20)
            self.window_view['table'].update(values=filtered_df.values.tolist(), auto_size_columns=True)

    def update_sgf_text(self):
        # Update the text element with SGF column content
        sgf_content = "\n".join(self.df["SGF"].astype(str))
        self.window["sgf_text"].update(sgf_content)

    def open_del_window(self):
        layout_del = [
            [sg.Text('Name of External shit:'), sg.InputText(key='Name')],
            [sg.Button('Find', key='find_btn'), sg.Button('Delete', key='delete_btn'),
             sg.Button('Cancel')],
            [sg.Table(values=self.df.tail(20).values.tolist(),
                      headings=self.headings, auto_size_columns=False,
                      justification='right', key='table')]
        ]

        self.window_del = sg.Window('Delete Window', layout_del, finalize=True)

        while True:
            event, values = self.window_del.read()

            if event == sg.WIN_CLOSED or event == 'Cancel':
                break
            elif event == 'find_btn':
                filtered_df = self.filter_data(values)
                self.window_del['table'].update(values=filtered_df.values.tolist())
            elif event == 'delete_btn':
                selected_row = values.get('table')
                if selected_row:
                    self.delete_record(selected_row[0], self.window_del)

        return self.window_del

    def delete_record(self, index, window_del):
        if index is not None and index in self.df.index:
            # Update the deletion_datetime column with the current date and time
            self.df.at[index, 'deletion_datetime'] = datetime.now()
            self.df.to_excel(self.database_path, index=False)
            sg.Popup('Record "deleted" successfully!')
            # Close the delete window
            window_del.close()
            # Update the view window table
            self.update_view_table()
        else:
            sg.PopupError('Error: Invalid index or index not found in DataFrame!')

            # Save the updated DataFrame to Excel
        self.df.to_excel(self.database_path, index=False)

    def run(self):
        while True:
            event, values = self.window.read()

            if event == sg.WIN_CLOSED or event == 'exit_btn':
                break
            elif event == 'input_btn':
                self.open_input_window()
            elif event == 'view_btn':
                self.open_view_window()
            elif event == 'del_btn':
                # Open the delete window
                self.window_del = self.open_del_window()
                # Get the event and values from the delete window
                event_del, values_del = self.window_del.read()
                if event_del == 'delete_btn':
                    selected_row = values_del.get('table')
                    if selected_row:
                        # Pass the window_del reference to the delete_record method
                        self.delete_record(selected_row, self.window_del)
                        sg.Popup('Record deleted successfully!')

        self.window.close()


if __name__ == '__main__':
    app = ExternalToolApp("C:\\ext_tool_DB\\external_db.xlsx")
    app.run()
