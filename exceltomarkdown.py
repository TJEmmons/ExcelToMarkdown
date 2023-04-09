import pandas as pd
from tkinter import *
from tkinter import messagebox
from pandastable import Table, config

class ExcelToMarkdownApp(Frame):
    def __init__(self, parent=None):
        self.parent = parent
        Frame.__init__(self)
        self.main = self.master
        self.main.geometry('800x600+200+100')
        self.main.title('Excel to Markdown Table Converter')
        
        self.init_ui()
    
    def init_ui(self):
        self.input_frame = Frame(self.main)
        self.input_frame.pack(side=TOP, fill=X)

        self.show_markdown_button = Button(self.input_frame, text='Show Markdown', command=self.show_markdown)
        self.show_markdown_button.pack(side=LEFT, padx=10, pady=10)

        self.output_frame = Frame(self.main)
        self.output_frame.pack(side=TOP, fill=BOTH, expand=1)

        self.init_table()
    
    def init_table(self):
        self.table_frame = Frame(self.output_frame)
        self.table_frame.pack(fill=BOTH, expand=1)

        self.dataframe = pd.DataFrame(index=range(100), columns=range(100))
        self.table = Table(self.table_frame, dataframe=self.dataframe,
                           showtoolbar=True, showstatusbar=True)
        self.table.show()
        
        options = {'colheadercolor': 'green', 'floatprecision': 5}
        config.apply_options(options, self.table)

    def show_markdown(self):
        self.dataframe = self.table.model.df.fillna("")
        markdown_table = self.dataframe.to_markdown(index=False)
        self.show_popup(markdown_table)


    def show_popup(self, text):
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
        
        text_widget.pack(expand=YES, fill=BOTH)
        
        scrollbar = Scrollbar(text_widget)
        scrollbar.pack(side=RIGHT, fill=Y)
        text_widget.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=text_widget.yview)
        
        # Copy the text to the clipboard
        self.clipboard_clear()
        self.clipboard_append(text)


if __name__ == '__main__':
    app = ExcelToMarkdownApp()
    app.mainloop()
