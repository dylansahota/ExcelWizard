from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from numpy import array_split
import pandas as pd
# Requires openpyxl to be installed on machine
# TO DO
# Re name all variables
# Re order label creation and grid commands
# Create open files and clear files to be one function rather than method
# Standardise file opening for the multiple file types
# Create Read Me
# Turn into Executable
# Fix Treeview first row and size of rows
# Add Horizontal Scrollbar to Treeview
# Fix label creation for blank row to one line rather than three lines
# Create function for turning file into dataframe
# Set default document directories to the same place

class MainApplication():
    def __init__(self, master):
        # Creating a frame in the main window - master which will take the argument of root when calling the class
        frame = Frame(master)
        frame.grid()

        # Row 1 - Blank Row
        self.SpaceLabel5 = Label(frame, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel5.grid(column = 1, row = 1)

        self.SpaceLabel6 = Label(frame, text = "", height = 1, width = 10, bg = "grey")
        self.SpaceLabel6.grid(column = 2, row = 1)

        self.SpaceLabel7 = Label(frame, text = "", height = 1, width = 45, bg = "grey")
        self.SpaceLabel7.grid(column = 3, row = 1)

        # Row 2 - Clean Button Row
        self.SpaceLabel4 = Label(frame, text = "", height = 3, width = 2, bg = "grey")
        self.SpaceLabel4.grid(column = 1, row = 2)

        self.RandomiserButton = Button(frame, text = "Clean", height = 3, width = 10, command = self.clean_window)
        self.RandomiserButton.grid(column = 2, row = 2)

        self.RandomiserLabel = Label(frame, text = "Removes any dodgy characters and trims all whitespaces", height = 3, width = 45, bg = "grey")
        self.RandomiserLabel.grid(column = 3, row = 2)

        # Row 3 - Blank Row
        self.SpaceLabel5 = Label(frame, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel5.grid(column = 1, row = 3)

        self.SpaceLabel6 = Label(frame, text = "", height = 1, width = 10, bg = "grey")
        self.SpaceLabel6.grid(column = 2, row = 3)

        self.SpaceLabel7 = Label(frame, text = "", height = 1, width = 45, bg = "grey")
        self.SpaceLabel7.grid(column = 3, row = 3)

        # Row 4 - Convert Button Row
        self.SpaceLabel8 = Label(frame, text = "", height = 3, width = 2, bg = "grey")
        self.SpaceLabel8.grid(column = 1, row = 4)

        self.ConvertButton = Button(frame, text = "Convert", height = 3, width = 10, command = self.convert_window)
        self.ConvertButton.grid(column = 2, row = 4)

        self.ConvertLabel = Label(frame, text = "Converts a file to XLSX or CSV", height = 3, width = 45, bg = "grey")
        self.ConvertLabel.grid(column = 3, row = 4)

        # Row 5 - Blank Row
        self.SpaceLabel9 = Label(frame, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel9.grid(column = 1, row = 5)

        self.SpaceLabel10 = Label(frame, text = "", height = 1, width = 10, bg = "grey")
        self.SpaceLabel10.grid(column = 2, row = 5)

        self.SpaceLabel11 = Label(frame, text = "", height = 1, width = 45, bg = "grey")
        self.SpaceLabel11.grid(column = 3, row = 5)

        #  Row 6 - Dedupe Button Row
        self.SpaceLabel12 = Label(frame, text = "", height = 3, width = 2, bg = "grey")
        self.SpaceLabel12.grid(column = 1, row = 6)

        self.DedupeButton = Button(frame, text = "Dedupe", height = 3, width = 10, command = self.dedupe_window)
        self.DedupeButton.grid(column = 2, row = 6)

        self.DedupeLabel = Label(frame, text = "Deduplicates a file based on values in a selected column", height = 3, width = 45, bg = "grey")
        self.DedupeLabel.grid(column = 3, row = 6)

        # Row 7 - Blank Row
        self.SpaceLabel13 = Label(frame, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel13.grid(column = 1, row = 7)

        self.SpaceLabel14 = Label(frame, text = "", height = 1, width = 10, bg = "grey")
        self.SpaceLabel14.grid(column = 2, row = 7)

        self.SpaceLabel15 = Label(frame, text = "", height = 1, width = 45, bg = "grey")
        self.SpaceLabel15.grid(column = 3, row = 7)

        # Row 8 - Edit Button Row
        self.SpaceLabel16 = Label(frame, text = "", height = 3, width = 2, bg = "grey")
        self.SpaceLabel16.grid(column = 1, row = 8)

        self.EditButton = Button(frame, text = "Edit", height = 3, width = 10, command = self.edit_window)
        self.EditButton.grid(column = 2, row = 8)

        self.EditLabel = Label(frame, text = "Adds/Removes selected columns in a file", height = 3, width = 45, bg = "grey")
        self.EditLabel.grid(column = 3, row = 8)

        # Row 9 - Blank Row
        self.SpaceLabel17 = Label(frame, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel17.grid(column = 1, row = 9)

        self.SpaceLabel18 = Label(frame, text = "", height = 1, width = 10, bg = "grey")
        self.SpaceLabel18.grid(column = 2, row = 9)

        self.SpaceLabel19 = Label(frame, text = "", height = 1, width = 45, bg = "grey")
        self.SpaceLabel19.grid(column = 3, row = 9)

        # Row 10 - Merge Button Row
        self.SpaceLabel20 = Label(frame, text = "", height = 3, width = 2, bg = "grey")
        self.SpaceLabel20.grid(column = 1, row = 10)

        self.MergeButton = Button(frame, text = "Merge", height = 3, width = 10, command = self.merge_window)
        self.MergeButton.grid(column = 2, row = 10)

        self.MergeLabel = Label(frame, text = "Merges multiple files together", height = 3, width = 45, bg = "grey")
        self.MergeLabel.grid(column = 3, row = 10)

        # Row 11 - Blank Row
        self.SpaceLabel21 = Label(frame, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel21.grid(column = 1, row = 11)

        self.SpaceLabel22 = Label(frame, text = "", height = 1, width = 10, bg = "grey")
        self.SpaceLabel22.grid(column = 2, row = 11)

        self.SpaceLabel23 = Label(frame, text = "", height = 1, width = 45, bg = "grey")
        self.SpaceLabel23.grid(column = 3, row = 11)

        # Row 12 - Preview Button Row
        self.SpaceLabel24 = Label(frame, text = "", height = 3, width = 2, bg = "grey")
        self.SpaceLabel24.grid(column = 1, row = 12)

        self.PreviewButton = Button(frame, text = "Preview", height = 3, width = 10, command = self.preview_window)
        self.PreviewButton.grid(column = 2, row = 12)

        self.PreviewLabel = Label(frame, text = "Preview the contents of a file", height = 3, width = 45, bg = "grey")
        self.PreviewLabel.grid(column = 3, row = 12)

        # Row 13 - Blank Row
        self.SpaceLabel25 = Label(frame, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel25.grid(column = 1, row = 13)

        self.SpaceLabel26 = Label(frame, text = "", height = 1, width = 10, bg = "grey")
        self.SpaceLabel26.grid(column = 2, row = 13)

        self.SpaceLabel27 = Label(frame, text = "", height = 1, width = 45, bg = "grey")
        self.SpaceLabel27.grid(column = 3, row = 13)

        # Row 14 - Randomiser Button Row
        self.SpaceLabel28 = Label(frame, text = "", height = 3, width = 2, bg = "grey")
        self.SpaceLabel28.grid(column = 1, row = 14)

        self.RandomiserButton = Button(frame, text = "Randomiser", height = 3, width = 10, command = self.randomise_window)
        self.RandomiserButton.grid(column = 2, row = 14)

        self.RandomiserLabel = Label(frame, text = "Randomise the order of a spreadsheet", height = 3, width = 45, bg = "grey")
        self.RandomiserLabel.grid(column = 3, row = 14)

        # Row 15 - Blank Row
        self.SpaceLabel30 = Label(frame, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel30.grid(column = 1, row = 15)

        self.SpaceLabel31 = Label(frame, text = "", height = 1, width = 10, bg = "grey")
        self.SpaceLabel31.grid(column = 2, row = 15)

        self.SpaceLabel32 = Label(frame, text = "", height = 1, width = 45, bg = "grey")
        self.SpaceLabel32.grid(column = 3, row = 15)

        # Row 16 - Sort Button Row
        self.SpaceLabel33 = Label(frame, text = "", height = 3, width = 2, bg = "grey")
        self.SpaceLabel33.grid(column = 1, row = 16)

        self.SortButton = Button(frame, text = "Sort", height = 3, width = 10, command = self.sort_window)
        self.SortButton.grid(column = 2, row = 16)

        self.SortLabel = Label(frame, text = "Sorts a file based on values in a selected column", height = 3, width = 45, bg = "grey")
        self.SortLabel.grid(column = 3, row = 16)

        # Row 17 - Blank Row
        self.SpaceLabel34 = Label(frame, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel34.grid(column = 1, row = 17)

        self.SpaceLabel35 = Label(frame, text = "", height = 1, width = 10, bg = "grey")
        self.SpaceLabel35.grid(column = 2, row = 17)

        self.SpaceLabel36 = Label(frame, text = "", height = 1, width = 45, bg = "grey")
        self.SpaceLabel36.grid(column = 3, row = 17)

        # Row 18 - Split Button Row
        self.SpaceLabel35 = Label(frame, text = "", height = 3, width = 2, bg = "grey")
        self.SpaceLabel35.grid(column = 1, row = 18)

        self.SplitButton = Button(frame, text = "Split", height = 3, width = 10, command = self.split_window)
        self.SplitButton.grid(column = 2, row = 18)

        self.SplitLabel = Label(frame, text = "Splits a file based on values in a selected column", height = 3, width = 45, bg = "grey")
        self.SplitLabel.grid(column = 3, row = 18)

        # Row 19 - Blank Row
        self.SpaceLabel36 = Label(frame, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel36.grid(column = 1, row = 19)

        self.SpaceLabel37 = Label(frame, text = "", height = 1, width = 10, bg = "grey")
        self.SpaceLabel37.grid(column = 2, row = 19)

        self.SpaceLabel38 = Label(frame, text = "", height = 1, width = 45, bg = "grey")
        self.SpaceLabel38.grid(column = 3, row = 19)

        # Row 20 - Separate Button Row
        self.SpaceLabel35 = Label(frame, text = "", height = 3, width = 2, bg = "grey")
        self.SpaceLabel35.grid(column = 1, row = 20)

        self.SeparateButton = Button(frame, text = "Seperate", height = 3, width = 10, command = self.seperate_window)
        self.SeparateButton.grid(column = 2, row = 20)

        self.SeparateLabel = Label(frame, text = "Separates multiple sheets into indivdual spreadsheets", height = 3, width = 45, bg = "grey")
        self.SeparateLabel.grid(column = 3, row = 20)

        # Row 21 - Blank Row
        self.SpaceLabel36 = Label(frame, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel36.grid(column = 1, row = 21)

        self.SpaceLabel37 = Label(frame, text = "", height = 1, width = 10, bg = "grey")
        self.SpaceLabel37.grid(column = 2, row = 21)

        self.SpaceLabel38 = Label(frame, text = "", height = 1, width = 45, bg = "grey")
        self.SpaceLabel38.grid(column = 3, row = 21)

        # Row 22 - Chunk Button Row
        self.SpaceLabel35 = Label(frame, text = "", height = 3, width = 2, bg = "grey")
        self.SpaceLabel35.grid(column = 1, row = 22)

        self.SeparateButton = Button(frame, text = "Chunk", height = 3, width = 10, command = self.chunk_window)
        self.SeparateButton.grid(column = 2, row = 22)

        self.SeparateLabel = Label(frame, text = "Splits a spreadsheet into smaller chunked spreadsheets", height = 3, width = 45, bg = "grey")
        self.SeparateLabel.grid(column = 3, row = 22)

        # Row 21 - Blank Row
        self.SpaceLabel36 = Label(frame, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel36.grid(column = 1, row = 23)

        self.SpaceLabel37 = Label(frame, text = "", height = 2, width = 10, bg = "grey")
        self.SpaceLabel37.grid(column = 2, row = 23)

        self.SpaceLabel38 = Label(frame, text = "", height = 2, width = 45, bg = "grey")
        self.SpaceLabel38.grid(column = 3, row = 23)

    # Method to call if button is picked Clean Window Class
    def clean_window(self):
        self.window = CleanWindow()

    def convert_window(self):
        self.window = ConvertWindow()

    def dedupe_window(self):
        self.window = DedupeWindow()

    def edit_window(self):
        self.window = EditWindow()

    def merge_window(self):
        self.window = MergeWindow()

    def preview_window(self):
        self.window = PreviewWindow()

    def randomise_window(self):
        self.window = RandomiseWindow()

    def sort_window(self):
        self.window = SortWindow()

    def split_window(self):
        self.window = SplitWindow()

    def chunk_window(self):
        self.window = ChunkWindow()

    def seperate_window(self):
        self.window = SeparateWindow()
        
# Class containing everything related to file cleaner
class CleanWindow():
    def __init__(self):
        self.top = Toplevel(bg="grey")
        self.top.title("Excel Wizard - File Cleaner")
        self.text = StringVar()
        self.text.set("")
        self.file_name = ""

        # Row 1 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 1, columnspan = 5)

        # Row 2 -  File Label Row
        self.SpaceLabel2 = Label(self.top, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel2.grid(column = 0, row = 2)

        self.FileExtLabel = Label(self.top, height = 2, width = 60, bg = "white", relief = "sunken", textvariable = self.text, anchor = "w")
        self.FileExtLabel.grid(column = 1, row = 2, columnspan = 3)

        self.SpaceLabel3 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel3.grid(column = 4, row = 2)

        # Row 3 -  Blank Row
        self.SpaceLabel4 = Label(self.top, text = "", height = 2, width = 64, bg = "grey")
        self.SpaceLabel4.grid(column = 0, row = 3, columnspan = 5)

        self.SpaceLabel5 = Label(self.top, text = "This will remove all non-ASCII characters and remove all whitespaces in a CSV/XLSX file", height = 2, width = 60, bg = "grey")
        self.SpaceLabel5.grid(column = 1, row = 3, columnspan = 3)

        self.SpaceLabel6 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel6.grid(column = 4, row = 2)

        # Row 4 -  Button Row
        self.SpaceLabel7 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel7.grid(column = 0, row = 4)

        self.SelectButton = Button(self.top, text = "Select File", height = 1, command = self.OpenFile)
        self.SelectButton.grid(column = 1, row = 4)

        self.ProcessButton = Button(self.top, text = "Process File", height = 1, command = self.ProcessFile)
        self.ProcessButton.grid(column = 2, row = 4)

        self.SplitButton = Button(self.top, text = "Clear File", height = 1, command = self.ClearFile)
        self.SplitButton.grid(column = 3, row = 4)

        self.SpaceLabel1 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel1.grid(column = 4, row = 4)

        # Row 5 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 5, columnspan = 5)

    # Method to select file and add it as a variable to be called when processing
    def OpenFile(self):
        self.text.set("")
        self.file_name = filedialog.askopenfilename(initialdir="/",title="Choose your file to be Cleaned")
        self.text.set(self.file_name)
        return self.file_name

    # Method to clear a previouly selected file
    def ClearFile(self):
        self.text.set("")
        self.file_name = ""

    # Method to clear out non-ASCII characters and all whitespaces from strings
    def ProcessFile(self):
        # Decides how to handle file depending on if it is a CSV or a XLSX otherwise throws an error message
        if self.file_name[-3:] == "csv":
            self.df = pd.read_csv(self.file_name)
            self.fileformat = "CSV"
        elif self.file_name[-4:] == "xlsx":
            if len(pd.ExcelFile(self.file_name).sheet_names) > 1:
                messagebox.showerror("Error!","This tool only supports Single Sheet Spreadsheets! Please Try Again")
                return False
            else:
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLSX"
        else:
            messagebox.showerror("Error!","This tool only supports CSV or XLSX files only! Please Try Again")

        # Looks for all ascii characters, and ignores/removes if there are any errors and then translates back to the normal characters
        self.df = self.df.applymap(lambda x: x.encode("ascii", errors="ignore").decode())
        # Looks for all columns which contain objects as datatypes
        self.df_obj = self.df.select_dtypes(['object'])
        # Removes all whitespaces from columns which contain strings, based on above variable
        self.df[self.df_obj.columns] = self.df_obj.apply(lambda x: x.str.strip())

        # Translates all files to CSV
        if self.fileformat == "CSV":
            self.df.to_csv(self.file_name[:-4]+"_Cleaned.csv", index = False)
        elif self.fileformat == "XLSX":
            self.df.to_csv(self.file_name[:-5]+"_Cleaned.csv", index = False)

        # Confirmation message after file has been cleaned
        messagebox.showinfo("Success!", "Your File has been Cleaned!")

# Class containing everything related to file cleaner
class ConvertWindow():
    def __init__(self):
        self.top = Toplevel(bg="grey")
        self.top.title("Excel Wizard - File Converter")
        self.text = StringVar()
        self.text.set("")
        self.file_name = ""

        # Row 1 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 1, columnspan = 7)

        # Row 2 -  File Label Row
        self.SpaceLabel2 = Label(self.top, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel2.grid(column = 0, row = 2)

        self.FileExtLabel = Label(self.top, height = 2, width = 60, bg = "white", relief = "sunken", textvariable = self.text, anchor = "w")
        self.FileExtLabel.grid(column = 1, row = 2, columnspan = 5)

        self.SpaceLabel3 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel3.grid(column = 6, row = 2)

        # Row 3 -  Blank Row
        self.SpaceLabel4 = Label(self.top, text = "", height = 2, width = 64, bg = "grey")
        self.SpaceLabel4.grid(column = 0, row = 3, columnspan = 7)


        # Row 4 - Description Row
        self.SpaceLabel5 = Label(self.top, text = "This will convert an excel spreadsheet to either CSV or XLSX", height = 2, width = 60, bg = "grey")
        self.SpaceLabel5.grid(column = 1, row = 3, columnspan = 5)

        # Row 4 -  Button Row
        self.SpaceLabel6 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel6.grid(column = 0, row = 4)

        self.SelectButton = Button(self.top, text = "Select File", height = 1, command = self.OpenFile)
        self.SelectButton.grid(column = 1, row = 4)

        self.ProcessButton = Button(self.top, text = "CSV Convert", height = 1, command = self.ProcessFileCSV)
        self.ProcessButton.grid(column = 2, row = 4)

        self.SpaceLabel7 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel7.grid(column = 3, row = 4)

        self.ProcessButton = Button(self.top, text = "XLSX Convert", height = 1, command = self.ProcessFileXLSX)
        self.ProcessButton.grid(column = 4, row = 4)

        self.SplitButton = Button(self.top, text = "Clear File", height = 1, command = self.ClearFile)
        self.SplitButton.grid(column = 5, row = 4)

        self.SpaceLabel8 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel8.grid(column = 6, row = 4)

        # Row 5 -  Blank Row
        self.SpaceLabel9 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel9.grid(column = 0, row = 5, columnspan = 7)

# Method to select file and add it as a variable to be called when processing
    def OpenFile(self):
        self.text.set("")
        self.file_name = filedialog.askopenfilename(initialdir="/",title="Choose your file to be Converted")
        self.text.set(self.file_name)
        return self.file_name

    # Method to clear a previouly selected file
    def ClearFile(self):
        self.text.set("")
        self.file_name = ""

    # Method to convert files to CSV/XLSX depending on the file input type
    def ProcessFileCSV(self):
        # Decides how to handle file depending on if it is a CSV or a XLSX otherwise throws an error message
        if self.file_name[-3:] == "csv" or self.file_name[-3:] == "txt":
            self.df = pd.read_csv(self.file_name)
            self.fileformat = "CSV"
        elif self.file_name[-4:] == "xlsx" or self.file_name[-4:] == "xlsm" or self.file_name[-4:] == "xlsb":
            if len(pd.ExcelFile(self.file_name).sheet_names) > 1:
                messagebox.showerror("Warning!","Converting this to CSV will cause you to lose any additional sheets! Proceed with Caution")
            else:
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLSX"
        elif self.file_name[-3:] == "xls":
            if len(pd.ExcelFile(self.file_name).sheet_names) > 1:
                messagebox.showerror("Warning!","Converting this to CSV will cause you to lose any additional sheets! Proceed with Caution")
            else:
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLS"
        else:
            messagebox.showerror("Error!","This tool only supports TXT, CSV or XLSX files only! Please Try Again")
            return False

        # Translates all files to CSV
        if self.fileformat == "CSV":
            self.df.to_csv(self.file_name[:-4]+".csv", index = False)
        elif self.fileformat == "XLSX":
            self.df.to_csv(self.file_name[:-5]+".csv", index = False)
        elif self.fileformat == "XLS":
            self.df.to_csv(self.file_name[:-4]+".csv", index = False)

        # Confirmation message after file has been cleaned
        messagebox.showinfo("Success!", "Your File has been Converted!")

        # Method to convert files to CSV/XLSX depending on the file input type
    def ProcessFileXLSX(self):
        # Decides how to handle file depending on if it is a CSV or a XLSX otherwise throws an error message
        if self.file_name[-3:] == "csv" or self.file_name[-3:] == "txt":
            self.df = pd.read_csv(self.file_name)
            self.fileformat = "CSV"
        elif self.file_name[-4:] == "xlsx" or self.file_name[-4:] == "xlsm" or self.file_name[-4:] == "xlsb":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLSX"
        elif self.file_name[-3:] == "xls":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLS"
        else:
            messagebox.showerror("Error!","This tool only supports TXT, CSV or XLSX files only! Please Try Again")
            return False

        # Translates all files to XLSX
        if self.fileformat == "CSV":
            self.df.to_excel(self.file_name[:-4]+".xlsx", index = False)
        elif self.fileformat == "XLSX":
            self.df.to_excel(self.file_name[:-5]+".xlsx", index = False)
        elif self.fileformat == "XLS":
            self.df.to_excel(self.file_name[:-4]+".xlsx", index = False)

        # Confirmation message after file has been cleaned
        messagebox.showinfo("Success!", "Your File has been Converted!")

# Class containing everything related to file deduper
class DedupeWindow():
    def __init__(self):
        self.top = Toplevel(bg="grey")
        self.top.title("Excel Wizard - File Deduper")
        self.text = StringVar()
        self.text.set("")
        self.file_name = ""

        # Row 1 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 1, columnspan = 5)

        # Row 2 -  File Label Row
        self.SpaceLabel2 = Label(self.top, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel2.grid(column = 0, row = 2)

        self.FileExtLabel = Label(self.top, height = 2, width = 60, bg = "white", relief = "sunken", textvariable = self.text, anchor = "w")
        self.FileExtLabel.grid(column = 1, row = 2, columnspan = 3)

        self.SpaceLabel3 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel3.grid(column = 4, row = 2)

        # Row 3 -  Blank Row
        self.SpaceLabel4 = Label(self.top, text = "", height = 2, width = 64, bg = "grey")
        self.SpaceLabel4.grid(column = 0, row = 3, columnspan = 5)

        self.SpaceLabel5 = Label(self.top, text = "This will deduplicate a file based on the values in a selected column", height = 2, width = 60, bg = "grey")
        self.SpaceLabel5.grid(column = 1, row = 3, columnspan = 3)

        self.SpaceLabel6 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel6.grid(column = 4, row = 2)

        # Row 4 -  Button Row
        self.SpaceLabel7 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel7.grid(column = 0, row = 4)

        self.SelectButton = Button(self.top, text = "Select File", height = 1, command = self.OpenFile)
        self.SelectButton.grid(column = 1, row = 4)

        self.ProcessButton = Button(self.top, text = "Select Column", height = 1, command = self.SelectColumn)
        self.ProcessButton.grid(column = 2, row = 4)

        self.SplitButton = Button(self.top, text = "Clear File", height = 1, command = self.ClearFile)
        self.SplitButton.grid(column = 3, row = 4)

        self.SpaceLabel1 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel1.grid(column = 4, row = 4)

        # Row 5 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 5, columnspan = 5)

    # Method to select file and add it as a variable to be called when processing
    def OpenFile(self):
        self.text.set("")
        self.file_name = filedialog.askopenfilename(initialdir="/documents/Excel Wizard Testing",title="Choose your file to be Deduplicated")
        self.text.set(self.file_name)
        return self.file_name

    # Method to clear a previouly selected file
    def ClearFile(self):
        self.text.set("")
        self.file_name = ""

    def SelectColumn(self):
        # Throws an error if there is no file selected in previous window
        if self.file_name == "":
            messagebox.showerror("Error!","You have not selected a File! Please try again")
            return False
        # Decides how to handle file depending on if it is a CSV or a XLSX otherwise throws an error message
        elif self.file_name[-3:] == "csv" or self.file_name[-3:] == "txt":
            self.df = pd.read_csv(self.file_name)
            self.fileformat = "CSV"
        elif self.file_name[-4:] == "xlsx" or self.file_name[-4:] == "xlsm" or self.file_name[-4:] == "xlsb":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLSX"
        elif self.file_name[-3:] == "xls":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLS"
        else:
            messagebox.showerror("Error!","This tool only supports TXT, CSV or XLSX files only! Please Try Again")
            return False

        # Creates a list based on the columns headers from the selected file
        column_list = list(self.df)

        # Creates a sub window to allow us to select column to dedupe file on
        self.subtop = Toplevel(bg="grey")
        self.subtop.geometry("350x275")
        self.subtop.title("Excel Wizard - Column Selector")

        # Creates listbox within sub-window which will contain column headers from selected file
        self.ListBox = Listbox(self.subtop, width = 35, height = 9)

        # loops through all items in column header list and adds them to listbox
        for item in column_list:
            self.ListBox.insert(END,item)

        # Adds listbox to sub-window
        self.ListBox.grid(pady = 15, padx = 17)

        # Inserts instruction comment
        self.CommentLabel = Label(self.subtop, text = "Please select the column you want to dedupe on", bg = "grey", pady = 10)
        self.CommentLabel.grid()

        # Adds dedupe button to sub-screen
        self.DedupeButton = Button(self.subtop, text = "Dedupe File", height = 1, command = self.DedupeProcessor)
        self.DedupeButton.grid()

    def DedupeProcessor(self):
        # Gets the selection from the list box, and enters it as a variable called field which is used in the dedupe action
        field = self.ListBox.get(ANCHOR)
        # Dedupes dataframe based on selected column from dataframe
        self.df.drop_duplicates(subset = field, keep = 'first', inplace = True)

        # Translates all files to the source file format
        if self.fileformat == "CSV":
            self.df.to_csv(self.file_name[:-4]+"_Deduped.csv", index = False)
        elif self.fileformat == "XLSX":
            self.df.to_excel(self.file_name[:-5]+"_Deduped.xlsx", index = False)
        elif self.fileformat == "XLS":
            self.df.to_excel(self.file_name[:-4]+"_Deduped.xlsx", index = False)

        # Confirmation message after file has been cleaned
        messagebox.showinfo("Success!", "Your File has been Deduped!")

# Class containing everything related to file deduper
class EditWindow():
    def __init__(self):
        self.top = Toplevel(bg="grey")
        self.top.title("Excel Wizard - File Editor")
        self.text = StringVar()
        self.text.set("")
        self.file_name = ""

        # Row 1 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 1, columnspan = 7)

        # Row 2 -  File Label Row
        self.SpaceLabel2 = Label(self.top, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel2.grid(column = 0, row = 2)

        self.FileExtLabel = Label(self.top, height = 2, width = 60, bg = "white", relief = "sunken", textvariable = self.text, anchor = "w")
        self.FileExtLabel.grid(column = 1, row = 2, columnspan = 5)

        self.SpaceLabel3 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel3.grid(column = 6, row = 2)

        # Row 3 -  Blank Row
        self.SpaceLabel4 = Label(self.top, text = "", height = 2, width = 64, bg = "grey")
        self.SpaceLabel4.grid(column = 0, row = 3, columnspan = 7)


        # Row 4 - Description Row
        self.SpaceLabel5 = Label(self.top, text = "This will Add or Remove columns from a file", height = 2, width = 60, bg = "grey")
        self.SpaceLabel5.grid(column = 1, row = 3, columnspan = 5)

        # Row 4 -  Button Row
        self.SpaceLabel6 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel6.grid(column = 0, row = 4)

        self.SelectButton = Button(self.top, text = "Select File", height = 1, command = self.OpenFile)
        self.SelectButton.grid(column = 1, row = 4)

        self.ProcessButton = Button(self.top, text = "Add Column", height = 1, command = self.AddColumn)
        self.ProcessButton.grid(column = 2, row = 4)

        self.SpaceLabel7 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel7.grid(column = 3, row = 4)

        self.ProcessButton = Button(self.top, text = "Remove Columns", height = 1, command = self.RemoveColumn)
        self.ProcessButton.grid(column = 4, row = 4)

        self.SplitButton = Button(self.top, text = "Clear File", height = 1, command = self.ClearFile)
        self.SplitButton.grid(column = 5, row = 4)

        self.SpaceLabel8 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel8.grid(column = 6, row = 4)

        # Row 5 - Blank Row
        self.SpaceLabel9 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel9.grid(column = 0, row = 5, columnspan = 7)

    # Method to select file and add it as a variable to be called when processing
    def OpenFile(self):
        self.text.set("")
        self.file_name = filedialog.askopenfilename(initialdir="/documents/Excel Wizard Testing",title="Choose your file to be Edited")
        self.text.set(self.file_name)
        return self.file_name

    # Method to clear a previouly selected file
    def ClearFile(self):
        self.text.set("")
        self.file_name = ""

    def AddColumn(self):
        # Throws an error if there is no file selected in previous window
        if self.file_name == "":
            messagebox.showerror("Error!","You have not selected a File! Please try again")
            return False
        # Decides how to handle file depending on if it is a CSV or a XLSX otherwise throws an error message
        elif self.file_name[-3:] == "csv" or self.file_name[-3:] == "txt":
            self.df = pd.read_csv(self.file_name)
            self.fileformat = "CSV"
        elif self.file_name[-4:] == "xlsx" or self.file_name[-4:] == "xlsm" or self.file_name[-4:] == "xlsb":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLSX"
        elif self.file_name[-3:] == "xls":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLS"
        else:
            messagebox.showerror("Error!","This tool only supports TXT, CSV or XLSX files only! Please Try Again")
            return False

        # Creates a list based on the columns headers from the selected file
        self.column_list = list(self.df)

        # Creates a sub window to allow us to select column to dedupe file on
        self.subtop = Toplevel(bg="grey")
        self.subtop.geometry("390x370")
        self.subtop.title("Excel Wizard - Column Selector")

        # Creates listbox within sub-window which will contain column headers from selected file
        self.ListBox = Listbox(self.subtop, width = 35, height = 9)

        # loops through all items in column header list and adds them to listbox
        for item in self.column_list:
            self.ListBox.insert(END,item)

        # Inserts instruction comment
        self.CommentLabel = Label(self.subtop, text = "Please select the column you want to insert a new column after", bg = "grey", pady = 5)
        self.CommentLabel.grid()

        # Adds listbox to sub-window
        self.ListBox.grid(padx = 17)

        # Inserts column header instruction
        self.CommentLabel = Label(self.subtop, text = "Please enter the name of the new column", bg = "grey", pady = 5)
        self.CommentLabel.grid()

        # Inserts Entry Box to allow insert of data into column addition
        self.ColumnEntryBox = Entry(self.subtop, width = 40)
        self.ColumnEntryBox.grid()

        # Inserts column header instruction
        self.CommentLabel = Label(self.subtop, text = "Please enter the default data for the new column", bg = "grey", pady = 5)
        self.CommentLabel.grid()

        # Inserts Entry Box to allow insert of data into column addition
        self.DataEntryBox = Entry(self.subtop, width = 40)
        self.DataEntryBox.grid()

        # Inserts instruction comment
        self.CommentLabel = Label(self.subtop, text = "", bg = "grey")
        self.CommentLabel.grid()

        # Adds dedupe button to sub-screen
        self.DedupeButton = Button(self.subtop, text = "Add Column", height = 1, command = self.AddColumnProcessor)
        self.DedupeButton.grid()

    def AddColumnProcessor(self):
        column_name = self.ColumnEntryBox.get()
        column_data = self.DataEntryBox.get()

        field = self.ListBox.get(ANCHOR)
        field_index = self.column_list.index(field)
        self.df.insert(field_index, column_name, column_data)

        # Translates all files to the source file format
        if self.fileformat == "CSV":
            self.df.to_csv(self.file_name[:-4]+"_Edited.csv", index = False)
        elif self.fileformat == "XLSX":
            self.df.to_excel(self.file_name[:-5]+"_Edited.xlsx", index = False)
        elif self.fileformat == "XLS":
            self.df.to_excel(self.file_name[:-4]+"_Edited.xlsx", index = False)

        # Confirmation message after file has been cleaned
        messagebox.showinfo("Success!", "Your Column has been Added!")

    def RemoveColumn(self):
        # Throws an error if there is no file selected in previous window
        if self.file_name == "":
            messagebox.showerror("Error!","You have not selected a File! Please try again")
            return False
        # Decides how to handle file depending on if it is a CSV or a XLSX otherwise throws an error message
        elif self.file_name[-3:] == "csv" or self.file_name[-3:] == "txt":
            self.df = pd.read_csv(self.file_name)
            self.fileformat = "CSV"
        elif self.file_name[-4:] == "xlsx" or self.file_name[-4:] == "xlsm" or self.file_name[-4:] == "xlsb":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLSX"
        elif self.file_name[-3:] == "xls":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLS"
        else:
            messagebox.showerror("Error!","This tool only supports TXT, CSV or XLSX files only! Please Try Again")
            return False

        # Creates a list based on the columns headers from the selected file
        self.column_list = list(self.df)

        # Creates a sub window to allow us to select column to dedupe file on
        self.subtop = Toplevel(bg="grey")
        self.subtop.geometry("350x250")
        self.subtop.title("Excel Wizard - Column Selector")

        # Creates listbox within sub-window which will contain column headers from selected file
        self.ListBox = Listbox(self.subtop, width = 35, height = 9, selectmode = MULTIPLE)

        # loops through all items in column header list and adds them to listbox
        for item in self.column_list:
            self.ListBox.insert(END,item)

        # Inserts instruction comment
        self.CommentLabel = Label(self.subtop, text = "Please select the columns you want to remove", bg = "grey", pady = 5)
        self.CommentLabel.grid()

        # Adds listbox to sub-window
        self.ListBox.grid(padx = 17)

        # Inserts instruction comment
        self.CommentLabel = Label(self.subtop, text = "", bg = "grey")
        self.CommentLabel.grid()

        # Adds dedupe button to sub-screen
        self.DedupeButton = Button(self.subtop, text = "Remove Column", height = 1, command = self.RemoveColumnProcessor)
        self.DedupeButton.grid()
    
    def RemoveColumnProcessor(self):

        self.column_indexes = list(self.ListBox.curselection())
        self.df.drop(self.df.columns[self.column_indexes], axis = 1, inplace = True)

        # Translates all files to the source file format
        if self.fileformat == "CSV":
            self.df.to_csv(self.file_name[:-4]+"_Edited.csv", index = False)
        elif self.fileformat == "XLSX":
            self.df.to_excel(self.file_name[:-5]+"_Edited.xlsx", index = False)
        elif self.fileformat == "XLS":
            self.df.to_excel(self.file_name[:-4]+"_Edited.xlsx", index = False)

        # Confirmation message after file has been cleaned
        messagebox.showinfo("Success!", "Your Columns have been Removed!")

class MergeWindow():
    def __init__(self):
        self.top = Toplevel(bg="grey")
        self.top.title("Excel Wizard - File Merger")
        self.text = StringVar()
        self.text.set("")
        self.text_string = ""
        self.file_name = ""

        # Row 1 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 1, columnspan = 5)

        # Row 2 -  File Label Row
        self.SpaceLabel2 = Label(self.top, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel2.grid(column = 0, row = 2)

        self.FileExtLabel = Label(self.top, height = 50, width = 60, bg = "white", relief = "sunken", textvariable = self.text, anchor = "nw", justify = LEFT)
        self.FileExtLabel.grid(column = 1, row = 2, columnspan = 3)

        self.SpaceLabel3 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel3.grid(column = 4, row = 2)

        # Row 3 -  Blank Row
        self.SpaceLabel4 = Label(self.top, text = "", height = 2, width = 64, bg = "grey")
        self.SpaceLabel4.grid(column = 0, row = 3, columnspan = 5)

        self.SpaceLabel5 = Label(self.top, text = "This will merge all selected files together into one final merged file", height = 2, width = 60, bg = "grey")
        self.SpaceLabel5.grid(column = 1, row = 3, columnspan = 3)

        self.SpaceLabel6 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel6.grid(column = 4, row = 2)

        # Row 4 -  Button Row
        self.SpaceLabel7 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel7.grid(column = 0, row = 4)

        self.SelectButton = Button(self.top, text = "Select Files", height = 1, command = self.OpenFile)
        self.SelectButton.grid(column = 1, row = 4)

        self.ProcessButton = Button(self.top, text = "Merge Files", height = 1, command = self.ProcessFile)
        self.ProcessButton.grid(column = 2, row = 4)

        self.SplitButton = Button(self.top, text = "Clear Files", height = 1, command = self.ClearFile)
        self.SplitButton.grid(column = 3, row = 4)

        self.SpaceLabel1 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel1.grid(column = 4, row = 4)

        # Row 5 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 5, columnspan = 5)

    # Method to select file and add it as a variable to be called when processing
    def OpenFile(self):
        self.text_string = filedialog.askopenfilenames(initialdir="/documents/Excel Wizard Testing",title="Choose your files to be Merged")
        # Loops through list of all selected files, and adds to the string variable with a new line between each file, for display purposes in the label widget
        for item in self.text_string:
            self.file_name = self.file_name+(item+"\n")
        self.text.set(self.file_name)

    # Method to clear a previouly selected file
    def ClearFile(self):
        self.text.set("")
        self.file_name = ""
        self.text_placeholder = ""
        self.text_string = ""
        self.filedirectory = ""

    # Method to clear out non-ASCII characters and all whitespaces from strings
    def ProcessFile(self):
        # Decides how to handle file depending on if it is a CSV or a XLSX otherwise throws an error message

        self.sourcestring = ""
        self.filedirectory = ""
        self.firstfile = ""
        self.df2 = pd.DataFrame()
        self.csv = False

        if len(self.text_string) < 2:
            messagebox.showerror("Error!","You haven't selected enough spreadsheets! Please Try Again")
            return False
        else:
            for item in self.text_string:
                if item[-3:] == "csv" or item[-3:] == "txt":
                    self.csv = True
                    self.df = pd.read_csv(item)
                elif item[-4:] == "xlsx" or item[-4:] == "xlsm" or item[-4:] == "xlsb":
                    if len(pd.ExcelFile(item).sheet_names) > 1:
                        messagebox.showerror("Error!","This tool only supports Single Sheet Spreadsheets! Please Try Again")
                        return False
                    else:
                        self.df = pd.read_excel(item)
                elif item[-3:] == "xls":
                    if len(pd.ExcelFile(item).sheet_names) > 1:
                        messagebox.showerror("Error!","This tool only supports Single Sheet Spreadsheets! Please Try Again")
                        return False
                    else:
                        self.df = pd.read_excel(item)
                else:
                    messagebox.showerror("Error!","This tool only supports CSV or XLSX files only! Please Try Again")

                self.sourcestring = item.split("/")
                self.df["source"] = self.sourcestring[-1]
                self.df2 = self.df2.append(self.df)
            
            self.df2.fillna('', inplace=True)
            self.firstfile = str(self.text_string[0])
            self.sourcestring = self.text_string[0].split("/")
            self.filedirectory = self.firstfile.replace(str(self.sourcestring[-1]),"")

            if self.csv == True:
                self.df2.to_csv(self.filedirectory+"/Mergelist.csv", index = False)
            else:
                self.df2.to_excel(self.filedirectory+"/Mergelist.xlsx", index = False)

            # Confirmation message after file has been cleaned
            messagebox.showinfo("Success!", "Your Files have been Merged!")

# Class containing everything related to file cleaner
class PreviewWindow():
    def __init__(self):
        self.top = Toplevel(bg="grey")
        self.top.title("Excel Wizard - File Preview")
        self.text = StringVar()
        self.text.set("")
        self.file_name = ""

        # Row 1 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 1, columnspan = 5)

        # Row 2 -  File Label Row
        self.SpaceLabel2 = Label(self.top, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel2.grid(column = 0, row = 2)

        self.FileExtLabel = Label(self.top, height = 2, width = 60, bg = "white", relief = "sunken", textvariable = self.text, anchor = "w")
        self.FileExtLabel.grid(column = 1, row = 2, columnspan = 3)

        self.SpaceLabel3 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel3.grid(column = 4, row = 2)

        # Row 3 -  Blank Row
        self.SpaceLabel4 = Label(self.top, text = "", height = 2, width = 64, bg = "grey")
        self.SpaceLabel4.grid(column = 0, row = 3, columnspan = 5)

        self.SpaceLabel5 = Label(self.top, text = "This will allow you to preview the contents of the first 100 rows of a file", height = 2, width = 60, bg = "grey")
        self.SpaceLabel5.grid(column = 1, row = 3, columnspan = 3)

        self.SpaceLabel6 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel6.grid(column = 4, row = 2)

        # Row 4 -  Button Row
        self.SpaceLabel7 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel7.grid(column = 0, row = 4)

        self.SelectButton = Button(self.top, text = "Select File", height = 1, command = self.OpenFile)
        self.SelectButton.grid(column = 1, row = 4)

        self.ProcessButton = Button(self.top, text = "Preview File", height = 1, command = self.PreviewFile)
        self.ProcessButton.grid(column = 2, row = 4)

        self.SplitButton = Button(self.top, text = "Clear File", height = 1, command = self.ClearFile)
        self.SplitButton.grid(column = 3, row = 4)

        self.SpaceLabel1 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel1.grid(column = 4, row = 4)

        # Row 5 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 5, columnspan = 5)

    # Method to select file and add it as a variable to be called when processing
    def OpenFile(self):
        self.text.set("")
        self.file_name = filedialog.askopenfilename(initialdir="/",title="Choose your file to be Cleaned")
        self.text.set(self.file_name)
        return self.file_name

    # Method to clear a previouly selected file
    def ClearFile(self):
        self.text.set("")
        self.file_name = ""

    # Method to clear out non-ASCII characters and all whitespaces from strings
    def PreviewFile(self):
        # Decides how to handle file depending on if it is a CSV or a XLSX otherwise throws an error message
        if self.file_name[-3:] == "csv":
            self.df = pd.read_csv(self.file_name)
            self.fileformat = "CSV"
        elif self.file_name[-4:] == "xlsx":
            if len(pd.ExcelFile(self.file_name).sheet_names) > 1:
                messagebox.showerror("Error!","This tool only supports Single Sheet Spreadsheets! Please Try Again")
                return False
            else:
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLSX"
        else:
            messagebox.showerror("Error!","This tool only supports CSV or XLSX files only! Please Try Again")
            return False

        # Checks length of dataframe, limits to 100 if the size is longer than 100
        if len(self.df) < 100:
            self.df = self.df
        else:
            self.df = self.df.head(100)

        # Reverses the order of the dataframe so it is in the same order as the original spreadsheet
        self.df = self.df.iloc[::-1]

        # Creates a list of the columns from the selected dataframe
        self.columns = list(self.df)

        # Creates a sub window which will contain preview of file
        self.subtop = Toplevel()
        self.subtop.title(self.file_name)

        # Creates Treeview
        self.treeview = ttk.Treeview(self.subtop)
        self.treeview.grid()

        # Inserts columns and data into treeview from dataframe
        self.treeview["columns"] = self.columns
        for item in self.columns:
            self.treeview.column(item, anchor = "w")
            self.treeview.heading(item, text = item, anchor = "w")
        
        for index, row in self.df.iterrows():
            self.treeview.insert("",0,text=index,values=list(row))

# Class containing everything related to file cleaner
class RandomiseWindow():
    def __init__(self):
        self.top = Toplevel(bg="grey")
        self.top.title("Excel Wizard - File Randomiser")
        self.text = StringVar()
        self.text.set("")
        self.file_name = ""

        # Row 1 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 1, columnspan = 5)

        # Row 2 -  File Label Row
        self.SpaceLabel2 = Label(self.top, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel2.grid(column = 0, row = 2)

        self.FileExtLabel = Label(self.top, height = 2, width = 60, bg = "white", relief = "sunken", textvariable = self.text, anchor = "w")
        self.FileExtLabel.grid(column = 1, row = 2, columnspan = 3)

        self.SpaceLabel3 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel3.grid(column = 4, row = 2)

        # Row 3 -  Blank Row
        self.SpaceLabel4 = Label(self.top, text = "", height = 2, width = 64, bg = "grey")
        self.SpaceLabel4.grid(column = 0, row = 3, columnspan = 5)

        self.SpaceLabel5 = Label(self.top, text = "This will randomise the order of a file", height = 2, width = 60, bg = "grey")
        self.SpaceLabel5.grid(column = 1, row = 3, columnspan = 3)

        self.SpaceLabel6 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel6.grid(column = 4, row = 2)

        # Row 4 -  Button Row
        self.SpaceLabel7 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel7.grid(column = 0, row = 4)

        self.SelectButton = Button(self.top, text = "Select File", height = 1, command = self.OpenFile)
        self.SelectButton.grid(column = 1, row = 4)

        self.ProcessButton = Button(self.top, text = "Randomise File", height = 1, command = self.RandomiseFile)
        self.ProcessButton.grid(column = 2, row = 4)

        self.SplitButton = Button(self.top, text = "Clear File", height = 1, command = self.ClearFile)
        self.SplitButton.grid(column = 3, row = 4)

        self.SpaceLabel1 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel1.grid(column = 4, row = 4)

        # Row 5 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 5, columnspan = 5)

    # Method to select file and add it as a variable to be called when processing
    def OpenFile(self):
        self.text.set("")
        self.file_name = filedialog.askopenfilename(initialdir="/",title="Choose your file to be Cleaned")
        self.text.set(self.file_name)
        return self.file_name

    # Method to clear a previouly selected file
    def ClearFile(self):
        self.text.set("")
        self.file_name = ""

    # Method to clear out non-ASCII characters and all whitespaces from strings
    def RandomiseFile(self):
        # Decides how to handle file depending on if it is a CSV or a XLSX otherwise throws an error message
        if self.file_name[-3:] == "csv":
            self.df = pd.read_csv(self.file_name)
            self.fileformat = "CSV"
        elif self.file_name[-4:] == "xlsx":
            if len(pd.ExcelFile(self.file_name).sheet_names) > 1:
                messagebox.showerror("Error!","This tool only supports Single Sheet Spreadsheets! Please Try Again")
                return False
            else:
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLSX"
        else:
            messagebox.showerror("Error!","This tool only supports CSV or XLSX files only! Please Try Again")
            return False

        # Randomises dataframe order
        self.df = self.df.sample(frac=1)

        # Translates all files to the source file format
        if self.fileformat == "CSV":
            self.df.to_csv(self.file_name, index = False)
        elif self.fileformat == "XLSX":
            self.df.to_excel(self.file_name, index = False)
        elif self.fileformat == "XLS":
            self.df.to_excel(self.file_name[:-4], index = False)

        # Confirmation message after file has been cleaned
        messagebox.showinfo("Success!", "Your File order has been randomised!")

# Class containing everything related to file deduper
class SortWindow():
    def __init__(self):
        self.top = Toplevel(bg="grey")
        self.top.title("Excel Wizard - File Sorter")
        self.text = StringVar()
        self.text.set("")
        self.file_name = ""

        # Row 1 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 1, columnspan = 5)

        # Row 2 -  File Label Row
        self.SpaceLabel2 = Label(self.top, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel2.grid(column = 0, row = 2)

        self.FileExtLabel = Label(self.top, height = 2, width = 60, bg = "white", relief = "sunken", textvariable = self.text, anchor = "w")
        self.FileExtLabel.grid(column = 1, row = 2, columnspan = 3)

        self.SpaceLabel3 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel3.grid(column = 4, row = 2)

        # Row 3 -  Blank Row
        self.SpaceLabel4 = Label(self.top, text = "", height = 2, width = 64, bg = "grey")
        self.SpaceLabel4.grid(column = 0, row = 3, columnspan = 5)

        self.SpaceLabel5 = Label(self.top, text = "This will sort a file based on the values in a column", height = 2, width = 60, bg = "grey")
        self.SpaceLabel5.grid(column = 1, row = 3, columnspan = 3)

        self.SpaceLabel6 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel6.grid(column = 4, row = 2)

        # Row 4 -  Button Row
        self.SpaceLabel7 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel7.grid(column = 0, row = 4)

        self.SelectButton = Button(self.top, text = "Select File", height = 1, command = self.OpenFile)
        self.SelectButton.grid(column = 1, row = 4)

        self.ProcessButton = Button(self.top, text = "Select Column", height = 1, command = self.SelectColumn)
        self.ProcessButton.grid(column = 2, row = 4)

        self.SplitButton = Button(self.top, text = "Clear File", height = 1, command = self.ClearFile)
        self.SplitButton.grid(column = 3, row = 4)

        self.SpaceLabel1 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel1.grid(column = 4, row = 4)

        # Row 5 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 5, columnspan = 5)

    # Method to select file and add it as a variable to be called when processing
    def OpenFile(self):
        self.text.set("")
        self.file_name = filedialog.askopenfilename(initialdir="/documents/Excel Wizard Testing",title="Choose your file to be Deduplicated")
        self.text.set(self.file_name)
        return self.file_name

    # Method to clear a previouly selected file
    def ClearFile(self):
        self.text.set("")
        self.file_name = ""

    def SelectColumn(self):
        # Throws an error if there is no file selected in previous window
        if self.file_name == "":
            messagebox.showerror("Error!","You have not selected a File! Please try again")
            return False
        # Decides how to handle file depending on if it is a CSV or a XLSX otherwise throws an error message
        elif self.file_name[-3:] == "csv" or self.file_name[-3:] == "txt":
            self.df = pd.read_csv(self.file_name)
            self.fileformat = "CSV"
        elif self.file_name[-4:] == "xlsx" or self.file_name[-4:] == "xlsm" or self.file_name[-4:] == "xlsb":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLSX"
        elif self.file_name[-3:] == "xls":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLS"
        else:
            messagebox.showerror("Error!","This tool only supports TXT, CSV or XLSX files only! Please Try Again")
            return False

        # Creates a list based on the columns headers from the selected file
        column_list = list(self.df)

        # Creates a sub window to allow us to select column to dedupe file on
        self.subtop = Toplevel(bg="grey")
        self.subtop.geometry("350x275")
        self.subtop.title("Excel Wizard - Column Selector")

        # Creates listbox within sub-window which will contain column headers from selected file
        self.ListBox = Listbox(self.subtop, width = 35, height = 9)

        # loops through all items in column header list and adds them to listbox
        for item in column_list:
            self.ListBox.insert(END,item)

        # Adds listbox to sub-window
        self.ListBox.grid(pady = 15, padx = 17)

        # Inserts instruction comment
        self.CommentLabel = Label(self.subtop, text = "Please select the column you want to sort on", bg = "grey", pady = 10)
        self.CommentLabel.grid()

        # Adds dedupe button to sub-screen
        self.DedupeButton = Button(self.subtop, text = "Sort File", height = 1, command = self.FileSorter)
        self.DedupeButton.grid()

    def FileSorter(self):
        # Gets the selection from the list box, and enters it as a variable called field which is used in the dedupe action
        field = self.ListBox.get(ANCHOR)
        # Sorts dataframe by the selected field
        self.df.sort_values(by = field, inplace=True)

        # Translates all files to the source file format
        if self.fileformat == "CSV":
            self.df.to_csv(self.file_name, index = False)
        elif self.fileformat == "XLSX":
            self.df.to_excel(self.file_name, index = False)
        elif self.fileformat == "XLS":
            self.df.to_excel(self.file_name, index = False)

        # Confirmation message after file has been cleaned
        messagebox.showinfo("Success!", "Your File has been Sorted!")

class SplitWindow():
    def __init__(self):
        self.top = Toplevel(bg="grey")
        self.top.title("Excel Wizard - File Splitter")
        self.text = StringVar()
        self.text.set("")
        self.text_string = ""
        self.file_name = ""

        # Row 1 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 1, columnspan = 5)

        # Row 2 -  File Label Row
        self.SpaceLabel2 = Label(self.top, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel2.grid(column = 0, row = 2)

        self.FileExtLabel = Label(self.top, height = 2, width = 60, bg = "white", relief = "sunken", textvariable = self.text, anchor = "w")
        self.FileExtLabel.grid(column = 1, row = 2, columnspan = 3)

        self.SpaceLabel3 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel3.grid(column = 4, row = 2)

        # Row 3 -  Blank Row
        self.SpaceLabel4 = Label(self.top, text = "", height = 2, width = 64, bg = "grey")
        self.SpaceLabel4.grid(column = 0, row = 3, columnspan = 5)

        self.SpaceLabel5 = Label(self.top, text = "This will split a file into smaller files based on the contents of a column", height = 2, width = 60, bg = "grey")
        self.SpaceLabel5.grid(column = 1, row = 3, columnspan = 3)

        self.SpaceLabel6 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel6.grid(column = 4, row = 2)

        # Row 4 -  Button Row
        self.SpaceLabel7 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel7.grid(column = 0, row = 4)

        self.SelectButton = Button(self.top, text = "Select File", height = 1, command = self.OpenFile)
        self.SelectButton.grid(column = 1, row = 4)

        self.ProcessButton = Button(self.top, text = "Select Column", height = 1, command = self.SelectColumn)
        self.ProcessButton.grid(column = 2, row = 4)

        self.SplitButton = Button(self.top, text = "Clear Files", height = 1, command = self.ClearFile)
        self.SplitButton.grid(column = 3, row = 4)

        self.SpaceLabel1 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel1.grid(column = 4, row = 4)

        # Row 5 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 5, columnspan = 5)

    # Method to select file and add it as a variable to be called when processing
    def OpenFile(self):
        self.text.set("")
        self.file_name = filedialog.askopenfilename(initialdir="/documents/Excel Wizard Testing",title="Choose your file to be Deduplicated")
        self.text.set(self.file_name)
        return self.file_name

    # Method to clear a previouly selected file
    def ClearFile(self):
        self.text.set("")
        self.file_name = ""
        self.text_placeholder = ""
        self.text_string = ""
        self.filedirectory = ""

    def SelectColumn(self):
        # Throws an error if there is no file selected in previous window
        if self.file_name == "":
            messagebox.showerror("Error!","You have not selected a File! Please try again")
            return False
        # Decides how to handle file depending on if it is a CSV or a XLSX otherwise throws an error message
        elif self.file_name[-3:] == "csv" or self.file_name[-3:] == "txt":
            self.df = pd.read_csv(self.file_name)
            self.fileformat = "CSV"
        elif self.file_name[-4:] == "xlsx" or self.file_name[-4:] == "xlsm" or self.file_name[-4:] == "xlsb":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLSX"
        elif self.file_name[-3:] == "xls":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLS"
        else:
            messagebox.showerror("Error!","This tool only supports TXT, CSV or XLSX files only! Please Try Again")
            return False

        # Creates a list based on the columns headers from the selected file
        column_list = list(self.df)

        # Creates a sub window to allow us to select column to dedupe file on
        self.subtop = Toplevel(bg="grey")
        self.subtop.geometry("350x275")
        self.subtop.title("Excel Wizard - Column Selector")

        # Creates listbox within sub-window which will contain column headers from selected file
        self.ListBox = Listbox(self.subtop, width = 35, height = 9)

        # loops through all items in column header list and adds them to listbox
        for item in column_list:
            self.ListBox.insert(END,item)

        # Adds listbox to sub-window
        self.ListBox.grid(pady = 15, padx = 17)

        # Inserts instruction comment
        self.CommentLabel = Label(self.subtop, text = "Please select the column you want to split file on", bg = "grey", pady = 10)
        self.CommentLabel.grid()

        # Adds dedupe button to sub-screen
        self.DedupeButton = Button(self.subtop, text = "Split File", height = 1, command = self.ProcessFile)
        self.DedupeButton.grid()

    # Method to clear out non-ASCII characters and all whitespaces from strings
    def ProcessFile(self):

        # Gets the selection from the list box, and enters it as a variable called field which is used in the split action
        field = self.ListBox.get(ANCHOR)

        # Splits file depending on the file format and the selected field from the list box
        if self.fileformat == "XLSX":
            for x in list(self.df[field].unique()):
                self.df[self.df[field] == x].to_excel(self.file_name[:-5]+"_{}.xlsx".format(x))
        elif self.fileformat == "XLS":
            for x in list(self.df[field].unique()):
                self.df[self.df[field] == x].to_excel(self.file_name[:-4]+"_{}.xlsx".format(x))
        elif self.fileformat == "CSV":
            for x in list(self.df[field].unique()):
                self.df[self.df[field] == x].to_excel(self.file_name[:-4]+"_{}.csv".format(x))

        # Confirmation message after file has been cleaned
        messagebox.showinfo("Success!", "Your Files have been Split!")

class ChunkWindow():
    def __init__(self):
        self.top = Toplevel(bg="grey")
        self.top.title("Excel Wizard - File Chunker")
        self.text = StringVar()
        self.text.set("")
        self.text_string = ""
        self.file_name = ""

        # Row 1 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 1, columnspan = 7)

        # Row 2 -  File Label Row
        self.SpaceLabel2 = Label(self.top, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel2.grid(column = 0, row = 2)

        self.FileExtLabel = Label(self.top, height = 2, width = 60, bg = "white", relief = "sunken", textvariable = self.text, anchor = "w")
        self.FileExtLabel.grid(column = 1, row = 2, columnspan = 5)

        self.SpaceLabel3 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel3.grid(column = 6, row = 2)

        # Row 3 -  Blank Row
        self.SpaceLabel4 = Label(self.top, text = "", height = 2, width = 64, bg = "grey")
        self.SpaceLabel4.grid(column = 0, row = 3, columnspan = 7)


        # Row 4 - Description Row
        self.SpaceLabel5 = Label(self.top, text = "This will chunk a file either by the number of chunks or the number of rows per chunk", height = 2, width = 60, bg = "grey")
        self.SpaceLabel5.grid(column = 1, row = 3, columnspan = 5)

        # Row 4 -  Button Row
        self.SpaceLabel6 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel6.grid(column = 0, row = 4)

        self.SelectButton = Button(self.top, text = "Select File", height = 1, command = self.OpenFile)
        self.SelectButton.grid(column = 1, row = 4)

        self.ProcessButton = Button(self.top, text = "Number of Chunks", height = 1, command = self.SelectChunks)
        self.ProcessButton.grid(column = 2, row = 4)

        self.SpaceLabel7 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel7.grid(column = 3, row = 4)

        self.ProcessButton = Button(self.top, text = "Number of Rows", height = 1, command = self.SelectRows)
        self.ProcessButton.grid(column = 4, row = 4)

        self.SplitButton = Button(self.top, text = "Clear File", height = 1, command = self.ClearFile)
        self.SplitButton.grid(column = 5, row = 4)

        self.SpaceLabel8 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel8.grid(column = 6, row = 4)

        # Row 5 - Blank Row
        self.SpaceLabel9 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel9.grid(column = 0, row = 5, columnspan = 7)

    # Method to select file and add it as a variable to be called when processing
    def OpenFile(self):
        self.text.set("")
        self.file_name = filedialog.askopenfilename(initialdir="/documents/Excel Wizard Testing",title="Choose your file to be Deduplicated")
        self.text.set(self.file_name)
        return self.file_name

    # Method to clear a previouly selected file
    def ClearFile(self):
        self.text.set("")
        self.file_name = ""
        self.text_placeholder = ""
        self.text_string = ""
        self.filedirectory = ""

    def SelectChunks(self):
        # Throws an error if there is no file selected in previous window
        if self.file_name == "":
            messagebox.showerror("Error!","You have not selected a File! Please try again")
            return False
        # Decides how to handle file depending on if it is a CSV or a XLSX otherwise throws an error message
        elif self.file_name[-3:] == "csv" or self.file_name[-3:] == "txt":
            self.df = pd.read_csv(self.file_name)
            self.fileformat = "CSV"
        elif self.file_name[-4:] == "xlsx" or self.file_name[-4:] == "xlsm" or self.file_name[-4:] == "xlsb":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLSX"
        elif self.file_name[-3:] == "xls":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLS"
        else:
            messagebox.showerror("Error!","This tool only supports TXT, CSV or XLSX files only! Please Try Again")
            return False

        # Creates a sub window to allow us to select column to dedupe file on
        self.subtop = Toplevel(bg="grey")
        self.subtop.geometry("395x130")
        self.subtop.title("Excel Wizard - Chunk Selector")

        self.BlankLabel = Label(self.subtop, text = "", bg = "grey", height = 1)
        self.BlankLabel.grid()

        self.Spinbox = Spinbox(self.subtop, from_=0,to=len(self.df), width = 30)
        self.Spinbox.grid()

        # Inserts instruction comment
        self.CommentLabel = Label(self.subtop, text = "Please select the number of chunks you want to split the file on", bg = "grey", pady = 10)
        self.CommentLabel.grid()

        # Adds dedupe button to sub-screen
        self.DedupeButton = Button(self.subtop, text = "Chunk File", height = 1, command = self.ProcessChunks)
        self.DedupeButton.grid()

    # Method to clear out non-ASCII characters and all whitespaces from strings
    def ProcessChunks(self):
        
        try:
            self.number_of_chunks = int(self.Spinbox.get())
        except:
            messagebox.showerror("Error!","You have not entered a number")
            return False

        if self.number_of_chunks > len(self.df):
            messagebox.showerror("Error!","You have selected more chunks than there are rows in your dataset")
            return False
        elif self.number_of_chunks == 0:
            messagebox.showerror("Error!","You can't split a file into 0 chunks")
            return False
        

        # Creates a variable to number the chunked files by
        self.file_number = 1

        # Splits the dataframe into chunks using Numpy array_split
        self.chunklist = array_split(self.df,self.number_of_chunks)

        # Sends each file in chunk list to the relative file format
        if self.fileformat == "XLSX":
            for x in self.chunklist:
                x.to_excel(self.file_name[:-5]+"_{}.xlsx".format(self.file_number))
                self.file_number += 1
        elif self.fileformat == "XLS":
            for x in self.chunklist:
                x.to_excel(self.file_name[:-4]+"_{}.xlsx".format(self.file_number))
                self.file_number += 1
        elif self.fileformat == "CSV":
            for x in self.chunklist:
                x.to_csv(self.file_name[:-4]+"_{}.csv".format(self.file_number))
                self.file_number += 1

        # Confirmation message after file has been cleaned
        messagebox.showinfo("Success!", "Your Files has been Chunked into {} Files!".format(self.number_of_chunks))
        self.file_number = 0

    def SelectRows(self):
        # Throws an error if there is no file selected in previous window
        if self.file_name == "":
            messagebox.showerror("Error!","You have not selected a File! Please try again")
            return False
        # Decides how to handle file depending on if it is a CSV or a XLSX otherwise throws an error message
        elif self.file_name[-3:] == "csv" or self.file_name[-3:] == "txt":
            self.df = pd.read_csv(self.file_name)
            self.fileformat = "CSV"
        elif self.file_name[-4:] == "xlsx" or self.file_name[-4:] == "xlsm" or self.file_name[-4:] == "xlsb":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLSX"
        elif self.file_name[-3:] == "xls":
                self.df = pd.read_excel(self.file_name)
                self.fileformat = "XLS"
        else:
            messagebox.showerror("Error!","This tool only supports TXT, CSV or XLSX files only! Please Try Again")
            return False

        # Creates a sub window to allow us to select column to dedupe file on
        self.subtop = Toplevel(bg="grey")
        self.subtop.geometry("335x130")
        self.subtop.title("Excel Wizard - Chunk Selector")

        self.BlankLabel = Label(self.subtop, text = "", bg = "grey", height = 1)
        self.BlankLabel.grid()

        self.Spinbox = Spinbox(self.subtop, from_=0, to = 1000000000, width = 30)
        self.Spinbox.grid()

        # Inserts instruction comment
        self.CommentLabel = Label(self.subtop, text = "Please select the number of rows you want per chunk", bg = "grey", pady = 10)
        self.CommentLabel.grid()

        # Adds dedupe button to sub-screen
        self.DedupeButton = Button(self.subtop, text = "Chunk File", height = 1, command = self.ProcessRows)
        self.DedupeButton.grid()

    # Method to clear out non-ASCII characters and all whitespaces from strings
    def ProcessRows(self):
        # Checks if entry is an integer, otherwise throws an error
        try:
            self.number_of_rows = int(self.Spinbox.get())
        except:
            messagebox.showerror("Error!","You have not entered a number")
            return False

        if self.number_of_rows > len(self.df):
            messagebox.showerror("Error!","You have selected more rows than there are rows in your dataset")
            return False
        elif self.number_of_rows == 0:
            messagebox.showerror("Error!","You can't split a file into 0 chunks")
            return False

        # Creates a variable to number the chunked files by
        self.file_number = 1
        self.chunksize = len(self.df)//self.number_of_rows

        # Splits the dataframe into chunks using Numpy array_split
        self.chunklist = array_split(self.df,self.chunksize)

        # Sends each file in chunk list to the relative file format
        if self.fileformat == "XLSX":
            for x in self.chunklist:
                x.to_excel(self.file_name[:-5]+"_{}.xlsx".format(self.file_number))
                self.file_number += 1
        elif self.fileformat == "XLS":
            for x in self.chunklist:
                x.to_excel(self.file_name[:-4]+"_{}.xlsx".format(self.file_number))
                self.file_number += 1
        elif self.fileformat == "CSV":
            for x in self.chunklist:
                x.to_csv(self.file_name[:-4]+"_{}.csv".format(self.file_number))
                self.file_number += 1

        # Confirmation message after file has been cleaned
        messagebox.showinfo("Success!", "Your Files has been Chunked into {} Files!".format(self.chunksize))
        self.file_number = 0

class SeparateWindow():
    def __init__(self):
        self.top = Toplevel(bg="grey")
        self.top.title("Excel Wizard - Sheet Separator")
        self.text = StringVar()
        self.text.set("")
        self.text_string = ""
        self.file_name = ""

        # Row 1 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 1, columnspan = 5)

        # Row 2 -  File Label Row
        self.SpaceLabel2 = Label(self.top, text = "", height = 1, width = 2, bg = "grey")
        self.SpaceLabel2.grid(column = 0, row = 2)

        self.FileExtLabel = Label(self.top, height = 2, width = 60, bg = "white", relief = "sunken", textvariable = self.text, anchor = "w")
        self.FileExtLabel.grid(column = 1, row = 2, columnspan = 3)

        self.SpaceLabel3 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel3.grid(column = 4, row = 2)

        # Row 3 -  Blank Row
        self.SpaceLabel4 = Label(self.top, text = "", height = 2, width = 64, bg = "grey")
        self.SpaceLabel4.grid(column = 0, row = 3, columnspan = 5)

        self.SpaceLabel5 = Label(self.top, text = "This will seperate multiple worksheets into seperate files", height = 2, width = 60, bg = "grey")
        self.SpaceLabel5.grid(column = 1, row = 3, columnspan = 3)

        self.SpaceLabel6 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel6.grid(column = 4, row = 2)

        # Row 4 -  Button Row
        self.SpaceLabel7 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel7.grid(column = 0, row = 4)

        self.SelectButton = Button(self.top, text = "Select File", height = 1, command = self.OpenFile)
        self.SelectButton.grid(column = 1, row = 4)

        self.ProcessButton = Button(self.top, text = "Separate File", height = 1, command = self.ProcessFile)
        self.ProcessButton.grid(column = 2, row = 4)

        self.SplitButton = Button(self.top, text = "Clear Files", height = 1, command = self.ClearFile)
        self.SplitButton.grid(column = 3, row = 4)

        self.SpaceLabel1 = Label(self.top, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel1.grid(column = 4, row = 4)

        # Row 5 -  Blank Row
        self.SpaceLabel1 = Label(self.top, text = "", height = 1, width = 64, bg = "grey")
        self.SpaceLabel1.grid(column = 0, row = 5, columnspan = 5)

    # Method to select file and add it as a variable to be called when processing
    def OpenFile(self):
        self.text.set("")
        self.file_name = filedialog.askopenfilename(initialdir="/documents/Excel Wizard Testing",title="Choose your file to be Deduplicated")
        self.text.set(self.file_name)
        return self.file_name

    # Method to clear a previouly selected file
    def ClearFile(self):
        self.text.set("")
        self.file_name = ""
        self.text_placeholder = ""
        self.text_string = ""
        self.filedirectory = ""

    def ProcessFile(self):
        # Throws an error if there is no file selected in previous window
        if self.file_name == "":
            messagebox.showerror("Error!","You have not selected a File! Please try again")
            return False
        # Decides how to handle file depending on if it is a CSV or a XLSX otherwise throws an error message
        elif self.file_name[-3:] == "csv" or self.file_name[-3:] == "txt":
            messagebox.showerror("Error!","This required multiple worksheets")
            return False
        elif self.file_name[-4:] == "xlsx" or self.file_name[-4:] == "xlsm" or self.file_name[-4:] == "xlsb":
            if len(pd.ExcelFile(self.file_name).sheet_names) < 2:
                messagebox.showerror("Error!","This tool only supports Multiple Sheet Spreadsheets! Please Try Again")
                return False
            else:
                self.filereader = pd.ExcelFile(self.file_name)
                # Reads all of the sheets in the path name as separate dataframes and writes them to excel
                for sheet in self.filereader.sheet_names:
                    self.df = pd.read_excel(self.file_name, sheet_name=sheet)
                    self.df.to_excel(self.file_name[:-5]+"_{}.xlsx".format(sheet),index = False)
        elif self.file_name[-3:] == "xls":
            if len(pd.ExcelFile(self.file_name).sheet_names) < 2:
                messagebox.showerror("Error!","This tool only supports Multiple Sheet Spreadsheets! Please Try Again")
                return False
            else:
                self.filereader = pd.ExcelFile(self.file_name)
                for sheet in self.filereader.sheet_names:
                    self.df = pd.read_excel(self.file_name, sheet_name=sheet)
                    self.df.to_excel(self.file_name[:-4]+"_{}.xlsx".format(sheet),index = False)
        else:
            messagebox.showerror("Error!","This tool only supports TXT, CSV or XLSX files only! Please Try Again")
            return False

        # Confirmation message after file has been cleaned
        messagebox.showinfo("Success!", "Your Sheets have been Separated!")
  
def main():
    root = Tk()
    root.title("Excel Wizard v1.0")
    app = MainApplication(root)
    root.mainloop()

if __name__ == "__main__":
    main()