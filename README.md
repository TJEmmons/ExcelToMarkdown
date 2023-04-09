# ExcelToMarkdown
Simple GUI to convert excel like data to github friendly markdown
Excel to Markdown Table Converter

Overview
This is a simple and easy-to-use Tkinter-based GUI application for converting Excel tables to Markdown tables. The application provides a pandas-based table with 100 rows and 100 columns, allowing users to paste in data from Excel. After pasting the data, the user can click on the 'Show Markdown' button to generate a Markdown version of the table, which is displayed in a popup window. The Markdown text is also copied to the clipboard automatically.

![image](https://user-images.githubusercontent.com/110789372/230797567-f27a53c0-1cb9-4887-a935-eeb890523dab.png)

This script requires the use of the pandastable tkinter object.  You can install this dependency using pip:

`pip install pandastable`

Once the application is running, follow these steps to convert an Excel table to Markdown format:

Copy the table from Excel.
Click on the first cell in the pandas table in the application.
Paste the table.
Click on the 'Show Markdown' button.
The Markdown version of the table will appear in a popup window, and it will also be copied to your clipboard.
You can now paste the Markdown table into any Markdown document or editor that supports the format.
