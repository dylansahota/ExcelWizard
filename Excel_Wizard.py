from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
# Requires openpyxl to be installed on machine

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

        self.DedupeButton = Button(frame, text = "Dedupe", height = 3, width = 10)#, command = self.MergeWindow)
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

        self.EditButton = Button(frame, text = "Edit", height = 3, width = 10)#, command = self.MergeWindow)
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

        self.MergeButton = Button(frame, text = "Merge", height = 3, width = 10)#, command = self.MergeWindow)
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

        self.PreviewButton = Button(frame, text = "Preview", height = 3, width = 10)#, command = self.MergeWindow)
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

        self.RandomiserButton = Button(frame, text = "Randomiser", height = 3, width = 10)#, command = self.MergeWindow)
        self.RandomiserButton.grid(column = 2, row = 14)

        self.RandomiserLabel = Label(frame, text = "Randomise the contents of a file based on a selected column", height = 3, width = 45, bg = "grey")
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

        self.SortButton = Button(frame, text = "Sort", height = 3, width = 10)#, command = self.MergeWindow)
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

        self.SplitButton = Button(frame, text = "Split", height = 3, width = 10)#, command = self.MergeWindow)
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

        self.SeparateButton = Button(frame, text = "Seperate", height = 3, width = 10)#, command = self.MergeWindow)
        self.SeparateButton.grid(column = 2, row = 20)

        self.SeparateLabel = Label(frame, text = "Separates multiple sheets into indivdual spreadsheets", height = 3, width = 45, bg = "grey")
        self.SeparateLabel.grid(column = 3, row = 20)

        # Row 21 - Blank Row
        self.SpaceLabel36 = Label(frame, text = "", height = 2, width = 2, bg = "grey")
        self.SpaceLabel36.grid(column = 1, row = 21)

        self.SpaceLabel37 = Label(frame, text = "", height = 2, width = 10, bg = "grey")
        self.SpaceLabel37.grid(column = 2, row = 21)

        self.SpaceLabel38 = Label(frame, text = "", height = 2, width = 45, bg = "grey")
        self.SpaceLabel38.grid(column = 3, row = 21)

    # Method to call if button is picked Clean Window Class
    def clean_window(self):
        self.window = CleanWindow()

    def convert_window(self):
        self.window = ConvertWindow()
        
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
            self.df.to_csv(self.file_name[:-4]+"_Cleaned.csv")
        elif self.fileformat == "XLSX":
            self.df.to_csv(self.file_name[:-5]+"_Cleaned.csv")

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
        self.file_name = filedialog.askopenfilename(initialdir="/",title="Choose your file to be Cleaned")
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
            self.df.to_csv(self.file_name[:-4]+".csv")
        elif self.fileformat == "XLSX":
            self.df.to_csv(self.file_name[:-5]+".csv")
        elif self.fileformat == "XLS":
            self.df.to_csv(self.file_name[:-4]+".csv")

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
            self.df.to_excel(self.file_name[:-4]+".xlsx")
        elif self.fileformat == "XLSX":
            self.df.to_excel(self.file_name[:-5]+".xlsx")
        elif self.fileformat == "XLS":
            self.df.to_excel(self.file_name[:-4]+".xlsx")

        # Confirmation message after file has been cleaned
        messagebox.showinfo("Success!", "Your File has been Converted!")

def main():
    root = Tk()
    root.title("Excel Wizard v1.0")
    app = MainApplication(root)
    root.mainloop()

if __name__ == "__main__":
    main()