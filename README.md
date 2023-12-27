External Tool Application
Overview
The External Tool Application is a Python desktop application designed to manage external tools, providing a user-friendly interface to input, view, and delete records. The application utilizes the PySimpleGUI library for the graphical user interface and Pandas for data manipulation.

Features
1. Input Records
Click the "Input" button to open the input window.
Enter details such as name, SGF (Serial Generated Field), and comments.
Validate the input based on specific criteria for each field.
Add records with a unique serial number and timestamp.
2. View Records
Click the "View" button to open the view window.
Filter records based on name and comments.
View a table displaying the last 20 records with columns for name, SGF, comments, and timestamp.
Dynamically update the table with filtering.
3. Delete Records
Click the "Delete" button to open the delete window.
Find records by name and comments.
Select records from the table and delete them.
Update the view window and exclude deleted records from the display.
4. Data Validation
Ensure data integrity with input validation for name, SGF, and comments.
Restrict SGF to a 6-digit number with specific criteria.
5. Automatic Column Sizing
Automatically adjust column widths in the view window for a better user experience.
6. Timestamps
Record the timestamp of record creation.
Record the timestamp of record deletion.
Usage
Run the application by executing the provided Python script.
Use the "Input," "View," and "Delete" buttons to navigate between functionalities.
Follow on-screen instructions for each operation, and interact with the input, view, and delete windows.
