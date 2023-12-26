import PySimpleGUI as sg
import pandas as pd
import openpyxl


class ExternalToolApp:
    def __init__(self, database_path):
        self.database_path = database_path
        self.df = pd.read_excel(database_path)
        self.headings = list(self.df.columns)

        sg.theme('LightBlue6')

        # Set the background color for the entire window
        window_background_color = '#74992e'  # Example color, you can change this

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

        self.window = sg.Window('External Tool', self.layout_enter, size=(400, 600), resizable=True, finalize=True,
                                background_color=window_background_color)
        self.window.set_min_size((400, 600))

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

        return (
                1 <= len(name) <= 20 and
                sgf.isdigit() and
                len(sgf) == 4 and
                1 <= int(sgf[:2]) <= 15 and
                sgf[2:] in {'20', '30', '40', '50', '60'} and
                1 <= len(comment) <= 40
        )

    def add_record(self, values):
        # Assuming 'Serial Number' is the first column in your DataFrame
        serial_number = len(self.df) + 1  # Incremented by 1 for the new row
        row = [serial_number, values['Name'], values['SGF'], values['Comment']]

        new_row = dict(zip(self.headings, row))
        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.df.to_excel(self.database_path, index=False)

    def open_view_window(self):
        layout_view = [
            [sg.Text('Name of External shit:'), sg.InputText(key='Name')],
            [sg.Text('SGF of External shit:'), sg.InputText(key='SGF')],
            [sg.Text('Comment of External shit:'), sg.InputText(key='Comment')],
            [sg.Button('Find', key='find_btn')],
            [sg.Table(values=self.df.tail(20).values.tolist(),
                      headings=self.headings, auto_size_columns=False,
                      justification='right', key='table')]
        ]

        window_view = sg.Window('View Window', layout_view)

        while True:
            event, values = window_view.read()

            if event == sg.WIN_CLOSED:
                break
            elif event == 'find_btn':
                filtered_df = self.filter_data(values)
                window_view['table'].update(values=filtered_df.values.tolist())
                self.update_sgf_text()  # Update SGF text after filtering

        window_view.close()

    def filter_data(self, values):
        name = values['Name']
        sgf = values['SGF']
        comment = values['Comment']

        query = ""
        if name:
            query += f"Name.str.contains('{name}')"

        if sgf:
            if query:
                query += " and "
            query += f"SGF.astype(str).str.contains('{sgf}') if 'SGF' in self.df.columns else False"

        if comment:
            if query:
                query += " and "
            query += f"Comment.str.contains('{comment}')"

        if query:
            return self.df.query(query)
        else:
            return self.df.tail(20)

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

        window_del = sg.Window('Delete Window', layout_del)

        while True:
            event, values = window_del.read()

            if event == sg.WIN_CLOSED or event == 'Cancel':
                break
            elif event == 'find_btn':
                filtered_df = self.filter_data(values)
                window_del['table'].update(values=filtered_df.values.tolist())
            elif event == 'delete_btn':
                selected_row = window_del['table'].get()
                if selected_row:
                    self.delete_record(selected_row[0])
                    sg.Popup('Record deleted successfully!')

        window_del.close()

    def delete_record(self, index):
        self.df = self.df.drop(index)
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
                self.open_del_window()

        self.window.close()


if __name__ == '__main__':
    app = ExternalToolApp("C:\\ext_tool_DB\\external_db.xlsx")
    app.run()
