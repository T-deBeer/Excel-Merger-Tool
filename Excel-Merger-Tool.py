import customtkinter
from tkinter import *
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import pandas as pd
import datetime as datetime
from pathlib import Path
import os
import openpyxl

window_height = 600
window_width = 1100
# colors
dark = "#264653"
dark_green = "#2A9D8F"
mustard = "#E9C46A"
orange = "#F4A261"
burnt = "#F4A261"

customtkinter.set_appearance_mode("Dark")
customtkinter.set_default_color_theme("dark-blue")

sheet_merge1 = pd.DataFrame()
sheet_merge2 = pd.DataFrame()


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # Functions
        def SelectFile1():
            filetypes = (
                ("Excel files", ".xlsx .xls"),
                ("All files", ".")
            )

            filename = fd.askopenfilename(
                title='Open a file',
                initialdir='/',
                filetypes=filetypes)

            self.txtFileOne.configure(state="normal")
            self.txtFileOne.delete(0, END)
            self.txtFileOne.insert(0, filename)
            self.txtFileOne.configure(state="readonly")

            if self.txtFileTwo.get() != "":
                common = GetCommonColumns(str(self.txtFileOne.get()),
                                          str(self.txtFileTwo.get()))
                self.cmbxCommonColumn.configure(values=common)

        def SelectFile2():
            filetypes = (
                ("Excel files", ".xlsx .xls"),
                ("All files", ".")
            )

            filename = fd.askopenfilename(
                title='Open a file',
                initialdir='/',
                filetypes=filetypes)

            self.txtFileTwo.configure(state="normal")
            self.txtFileTwo.delete(0, END)
            self.txtFileTwo.insert(0, filename)
            self.txtFileTwo.configure(state="readonly")

            if self.txtFileOne.get() != "":
                common = GetCommonColumns(str(self.txtFileOne.get()),
                                          str(self.txtFileTwo.get()))
                self.cmbxCommonColumn.configure(values=common)

        def SelectColumnFile():
            columns = []
            filetypes = (
                ("Excel files", ".xlsx .xls"),
                ("All files", ".")
            )

            filename = fd.askopenfilename(
                title='Open a file',
                initialdir='/',
                filetypes=filetypes)

            self.txtSelectFile.configure(state="normal")
            self.txtSelectFile.delete(0, END)
            self.txtSelectFile.insert(0, filename)
            self.txtSelectFile.configure(state="readonly")

            df = pd.read_excel(filename, sheet_name=0)
            for col in df.columns:
                columns.append(col)

            self.cmbxColumn1.configure(values=columns)

        def GetCommonColumns(file1, file2):
            sheet1_columns = []
            sheet2_columns = []
            common_columns = []
            sheet_merge1 = pd.read_excel(
                file1, sheet_name=0)
            sheet_merge2 = pd.read_excel(
                file2, sheet_name=0)
            for i in sheet_merge1:
                sheet1_columns.append(i)
            for i in sheet_merge2:
                sheet2_columns.append(i)

            if len(sheet1_columns) >= len(sheet2_columns):
                for i in sheet1_columns:
                    if i in sheet2_columns:
                        common_columns.append(i)
            else:
                for i in sheet2_columns:
                    if i in sheet1_columns:
                        common_columns.append(i)

            return common_columns

        def MergeSheets(header):
            self.lblMessage = customtkinter.CTkLabel(
                sheet_merging_tab, text="Merging Sheets....", font=customtkinter.CTkFont(size=15), text_color=mustard)
            self.lblMessage .grid(row=4, column=1, pady=10)
            sheet_merging_tab.update()

            sheet_merge1 = pd.read_excel(
                str(self.txtFileOne.get()), sheet_name=0)
            sheet_merge2 = pd.read_excel(
                str(self.txtFileTwo.get()), sheet_name=0)

            df = sheet_merge1.merge(sheet_merge2, on=header, how='outer')

            df.loc[df[header].duplicated(), header] = pd.NA

            today = datetime.datetime.now()
            today = today.strftime("%Y-%m-%d_%H-%M")

            path = f'{Path.cwd()}\VHT Merged_{today}.xlsx'

            # export new dataframe to excel
            df.to_excel(f'VHT Merged_{today}.xlsx', index=False)

            showinfo(
                title='New Sheet Made',
                message=f"File can be found at:\n{Path.cwd()}\VHT Merged_{today}.xlsx"
            )
            os.startfile(path)

            self.txtFileOne.configure(state="normal")
            self.txtFileOne.delete(0, END)
            self.txtFileOne.configure(state="readonly")
            self.txtFileTwo.configure(state="normal")
            self.txtFileTwo.delete(0, END)
            self.txtFileTwo.configure(state="readonly")
            self.cmbxCommonColumn.configure(values=[])
            self.cmbxCommonColumn.set("Common Columns")

            self.lblMessage.destroy()

        def SelectColumn1(selected_column):
            columns = []

            df = pd.read_excel(str(self.txtSelectFile.get()), sheet_name=0)
            for col in df.columns:
                if col != selected_column and df[col].dtypes == df[selected_column].dtypes:
                    columns.append(col)

            self.cmbxColumn2.configure(state="normal")
            self.cmbxColumn2.configure(values=columns)

        def SelectColumn2(second_column):
            value = []
            value.append(str(self.cmbxColumn1.get()))
            value.append(second_column)

            self.cmbxColumn.configure(state="normal")
            self.cmbxColumn.configure(values=value)

        def MergeColumns(precedence):
            columns = []

            column_one = str(self.cmbxColumn1.get())
            column_two = str(self.cmbxColumn2.get())

            df = pd.read_excel(str(self.txtSelectFile.get()), sheet_name=0)
            for col in df.columns:
                columns.append(col)

            new_column = ""
            dialog = customtkinter.CTkInputDialog(
                text="Type the name of the joint column:", title="Joint Column Name", fg_color=dark, button_fg_color=dark_green, button_hover_color=mustard, entry_border_color=dark_green, entry_fg_color=dark)

            input = dialog.get_input()
            print(input)
            if input in columns:
                new_column = input + "1"
            else:
                new_column = input

            new_values = []
            for i in df.index:
                if df[column_one][i] is None:
                    new_values.append(df[column_two][i])
                else:
                    if df[column_two][i] is None:
                        new_values.append(df[column_one][i])
                    else:
                        new_values.append(df[precedence][i])

            df[new_column] = new_values

            today = datetime.datetime.now()
            today = today.strftime("%Y-%m-%d_%H-%M")

            wb = openpyxl.load_workbook(
                str(self.txtSelectFile.get()))
            sheets = wb.sheetnames
            wb.remove(wb[sheets[0]])

            df.to_excel(str(self.txtSelectFile.get()),
                        sheet_name=f"Updated {today}")

            showinfo(
                title='Columns Joined',
                message=f"Columns have been joined successfully")

            self.txtSelectFile.configure(state="normal")
            self.txtSelectFile.delete(0, END)
            self.txtSelectFile.configure(state="readonly")

            self.cmbxColumn1.configure(state="normal")
            self.cmbxColumn1.configure(values=[])
            self.cmbxColumn1.set("First Column")

            self.cmbxColumn2.configure(state="normal")
            self.cmbxColumn2.configure(values=[])
            self.cmbxColumn2.set("Second Column")

            self.cmbxColumn.configure(state="normal")
            self.cmbxColumn.configure(values=[])
            self.cmbxColumn.set("Precedence Column")

        # configure window
        self.title("Excel Merging Tool")
        self.resizable(0, 0)
        self.iconbitmap(".\Images\Logo.ico")

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        x_cordinate = int((screen_width/2) - (window_width/2))
        y_cordinate = int((screen_height/2) - (window_height/2))

        self.geometry("{}x{}+{}+{}".format(window_width,
                      window_height, x_cordinate, y_cordinate))

        # Frame elements
        self.tabview = customtkinter.CTkTabview(
            self, fg_color=dark, border_width=2, corner_radius=5, border_color=dark_green, segmented_button_fg_color=dark, segmented_button_selected_color=dark_green, segmented_button_unselected_hover_color=dark_green, segmented_button_unselected_color=dark, segmented_button_selected_hover_color=dark_green)
        self.tabview.pack(side=LEFT)
        self.tabview.configure(width=1100, height=600)

        # Merge sheets tab
        sheet_merging_tab = self.tabview.add("Merge Sheets")
        # Tab configuration
        sheet_merging_tab.columnconfigure((0, 1, 2), weight=1)
        sheet_merging_tab.rowconfigure((0, 1, 2), weight=0)
        # Tab components
        self.lblRequirements = customtkinter.CTkTextbox(
            sheet_merging_tab, font=customtkinter.CTkFont(size=18), fg_color="transparent", width=590, height=100)
        self.lblRequirements.insert(
            "0.0", "Merge the first sheet of two seperate Excel files. The sheets must have atleast one column heading in common.")
        self.lblRequirements.configure(state="disabled")
        self.lblRequirements.grid(row=0, column=0, columnspan=3)
        # File One components
        self.lblFileOne = customtkinter.CTkLabel(
            sheet_merging_tab, text="File One", font=customtkinter.CTkFont(size=15))
        self.lblFileOne.grid(row=1, column=0, pady=10)

        self.txtFileOne = customtkinter.CTkEntry(sheet_merging_tab,
                                                 font=customtkinter.CTkFont(size=15), width=580, fg_color=dark, border_color=dark_green)
        self.txtFileOne.grid(row=1, column=1, padx=10, pady=10)

        self.btnFileOne = customtkinter.CTkButton(
            sheet_merging_tab, text="File One", fg_color=dark, hover_color=dark_green, border_width=2, corner_radius=5, border_color=dark_green, font=customtkinter.CTkFont(size=15), command=SelectFile1)
        self.btnFileOne.grid(row=1, column=2, padx=25, pady=10)
        # File Two components
        self.lblFileTwo = customtkinter.CTkLabel(
            sheet_merging_tab, text="File Two", font=customtkinter.CTkFont(size=15))
        self.lblFileTwo .grid(row=2, column=0, pady=10)

        self.txtFileTwo = customtkinter.CTkEntry(sheet_merging_tab,
                                                 font=customtkinter.CTkFont(size=15), width=580, fg_color=dark, border_color=dark_green)
        self.txtFileTwo .grid(row=2, column=1, padx=10, pady=10)

        self.btnFileTwo = customtkinter.CTkButton(
            sheet_merging_tab, text="File Two", fg_color=dark, hover_color=dark_green, border_width=2, corner_radius=5, border_color=dark_green, font=customtkinter.CTkFont(size=15), command=SelectFile2)
        self.btnFileTwo .grid(row=2, column=2, padx=25, pady=10)
        # Common column selction
        self.lblCommonColumn = customtkinter.CTkLabel(
            sheet_merging_tab, text="Select Common Column", font=customtkinter.CTkFont(size=15))
        self.lblCommonColumn .grid(row=3, column=0, pady=10)
        self.cmbxCommonColumn = customtkinter.CTkComboBox(
            sheet_merging_tab, values=[], border_color=dark_green, dropdown_hover_color=dark_green, fg_color=dark, corner_radius=5, width=250, command=MergeSheets)
        self.cmbxCommonColumn.set("Common Column")
        self.cmbxCommonColumn .grid(row=3, column=1, padx=25, pady=10)
        # Merge column tabs
        column_merging_tab = self.tabview.add("Merge Columns")
        # Tab configuration
        column_merging_tab.columnconfigure((0, 1, 2), weight=1)
        column_merging_tab.rowconfigure((0, 1, 2), weight=0)
        # Tab Components
        self.lblDesc = customtkinter.CTkTextbox(
            column_merging_tab, font=customtkinter.CTkFont(size=18), fg_color="transparent", width=580, height=100)
        self.lblDesc.insert(
            "0.0", "Merge seperate columns in a sheet into one column. If both columns have values then the precedence column's value will be used.")
        self.lblDesc.configure(state="disabled")
        self.lblDesc.grid(row=0, column=0, columnspan=3)

        self.lblSelectFile = customtkinter.CTkLabel(
            column_merging_tab, text="Select File", font=customtkinter.CTkFont(size=15))
        self.lblSelectFile .grid(row=1, column=0, pady=10)

        self.txtSelectFile = customtkinter.CTkEntry(column_merging_tab,
                                                    font=customtkinter.CTkFont(size=15), width=580, fg_color=dark, border_color=dark_green)
        self.txtSelectFile .grid(row=1, column=1, padx=10, pady=10)

        self.btnSelectFile = customtkinter.CTkButton(
            column_merging_tab, text="Select File", fg_color=dark, hover_color=dark_green, border_width=2, corner_radius=5, border_color=dark_green, font=customtkinter.CTkFont(size=15), command=SelectColumnFile)
        self.btnSelectFile .grid(row=1, column=2, padx=25, pady=10)

        self.cmbxColumn1 = customtkinter.CTkComboBox(
            column_merging_tab, values=[], border_color=dark_green, dropdown_hover_color=dark_green, fg_color=dark, corner_radius=5, width=300, command=SelectColumn1)
        self.cmbxColumn1.set("First Column")
        self.cmbxColumn1 .grid(row=2, column=1, padx=5, pady=10)

        self.cmbxColumn2 = customtkinter.CTkComboBox(
            column_merging_tab, values=[], border_color=dark_green, dropdown_hover_color=dark_green, fg_color=dark, corner_radius=5, width=300, command=SelectColumn2)
        self.cmbxColumn2.set("Second Column")
        self.cmbxColumn2.configure(state="disabled")
        self.cmbxColumn2 .grid(row=3, column=1, padx=5, pady=10)

        self.cmbxColumn = customtkinter.CTkComboBox(
            column_merging_tab, values=[], border_color=dark_green, dropdown_hover_color=dark_green, fg_color=dark, corner_radius=5, width=300, command=MergeColumns)
        self.cmbxColumn.set("Precedence Column")
        self.cmbxColumn.configure(state="disabled")
        self.cmbxColumn .grid(row=4, column=1, padx=5, pady=10)


if __name__ == "__main__":
    app = App()
    app.mainloop()
