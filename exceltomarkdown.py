"""
Excel to Markdown Table Converter

This script creates a simple Tkinter-based GUI application to convert Excel tables
to Markdown tables. The application provides a pandas-based table with 100 rows and
100 columns, allowing users to paste in data from Excel. The user can then click on
the 'Show Markdown' button to generate a Markdown version of the table, which is
displayed in a popup window. The Markdown text is also copied to the clipboard
automatically.
"""

import pandas as pd
from tkinter import *
from tkinter import messagebox
from pandastable import Table, config

# Create a class for the GUI application
class ExcelToMarkdownApp(Frame):
    def __init__(self, parent=None):
        self.parent = parent
        Frame.__init__(self)
        self.main = self.master
        self.main.geometry('800x600+200+100')
        self.main.title('Excel to Markdown Table Converter')
        
        self.init_ui()  # Call the init_ui function upon initialization
    
    def init_ui(self):
        # Create input frame
        self.input_frame = Frame(self.main)
        self.input_frame.pack(side=TOP, fill=X)

        # Create show markdown button
        self.show_markdown_button = Button(self.input_frame, text='Show Markdown', command=self.show_markdown)
        self.show_markdown_button.pack(side=LEFT, padx=10, pady=10)

        # Create output frame
        self.output_frame = Frame(self.main)
        self.output_frame.pack(side=TOP, fill=BOTH, expand=1)

        self.init_table()  # Call the init_table function to create table
    
    def init_table(self):
        # Create frame for table
        self.table_frame = Frame(self.output_frame)
        self.table_frame.pack(fill=BOTH, expand=1)

        # Create dataframe to display in table
        self.dataframe = pd.DataFrame(index=range(100), columns=range(100))
        self.table = Table(self.table_frame, dataframe=self.dataframe,
                           showtoolbar=True, showstatusbar=True)
        self.table.show()
        
        # Set options for the table
        options = {'colheadercolor': 'green', 'floatprecision': 5}
        config.apply_options(options, self.table)

    def show_markdown(self):
        # Convert the data in the table to a markdown table
        self.dataframe = self.table.model.df.fillna("")
        markdown_table = self.dataframe.to_markdown(index=False)
        self.show_popup(markdown_table)  # Show the markdown table in a popup window


    def show_popup(self, text):
        # Create a popup window to display the markdown table
        popup = Toplevel()
        popup.title("Markdown Table")
        text_widget = Text(popup, wrap=WORD)
        
        # Insert the text into the Text widget
        text_widget.insert(INSERT, text)
        
        # Get the bounding box of the last character in the Text widget
        bbox = text_widget.bbox(END)
        if bbox:
            width = bbox[0] + bbox[2] + 20  # Add some extra space for the scrollbar
            height = bbox[1] + bbox[3] + 20  # Add some extra space for the title bar
            popup.geometry(f"{width}x{height}")
        
        # Pack the Text widget and scrollbar
        text_widget.pack(expand=YES, fill=BOTH)
        scrollbar = Scrollbar(text_widget)
        scrollbar.pack(side=RIGHT, fill=Y)
        text_widget.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=text_widget.yview)
        
        # Copy the text to the clipboard
        self.clipboard_clear()
        self.clipboard_append(text)

# Run the GUI application
if __name__ == '__main__':
    app = ExcelToMarkdownApp()
    app.mainloop()
