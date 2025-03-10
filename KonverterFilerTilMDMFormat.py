import openpyxl
import openpyxl.workbook
from tkinter import filedialog, messagebox
import json
import os
from pathlib import Path


class transformExcel:

    def getInputData(self):
        self.RowsConfig = ""

        try:
            with open("configFK.json", "r") as f:
                try:
                    self.RowsConfig = json.load(f)
                except ValueError as e:
                    print(e, "\nConfig filen virker å være korrupt")
        except FileNotFoundError as e:
            print(e, "\nkunne ikke finne config filen")

        self.filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")], title="Select file"
        )

    def variableSetup(self):

        if self.RowsConfig and self.filename:
            self.StartRow = self.RowsConfig.get("startrow") or 1
            self.outPutRow = self.RowsConfig.get("outPutRow") or 1
            self.PrimaryKeyColumn = self.RowsConfig.get("PrimarKeyColumn") or "A"

            self.TargetRowOffset = range(len(self.RowsConfig.get("Rows"))) or 1
            self.itemOfset = len(self.RowsConfig.get("Rows")) or 1

            self.wb = openpyxl.load_workbook(filename=self.filename)
            sheet = self.RowsConfig.get("readSheet") or self.wb.active.title
            try:
                self.ws = self.wb[sheet]
            except:
                self.ws = None

    def setupNewFile(self):
        self.new_wb = openpyxl.load_workbook(r"sdfa.xlsx") # legg inn path til mal Thor Isum
        newsheet = self.RowsConfig.get("OutSheet") or self.new_wb.active.title
        #if not newsheet in self.new_wb.worksheets:
        #    print(newsheet)     
        #    self.new_wb.create_sheet(newsheet, 0)
        self.new_ws = self.new_wb[newsheet]

    def readAndWriteExcel(self):
        for i, self.rowSelected in enumerate(self.RowsConfig["Rows"]):
            for row in range(10_000):

                self.deleteIfRowFailes = [] # stores column and row so we can delete if flagg skipRowIfMissing is triggered
                offset = (self.itemOfset * row) + self.skippedRows
                self.NewWorkBookCurrentRow = self.outPutRow + self.TargetRowOffset[i] + offset

                if not self.ws[self.PrimaryKeyColumn + str(self.StartRow + row)].value:
                    break  # Stoppes if PrimaryKeyColumn is empty
                for col in self.rowSelected.keys():
                    self.skipRowIfMissingFlagg = self.rowSelected[col].get("skipRowIfMissing")
                    valueToPlace = (
                        self.rowSelected[col].get("Value")
                        or self.ws[col + str(self.StartRow + row)].value
                    ) # try to set default value, if not checks cell
                    if valueToPlace:
                        self.new_ws[
                            self.rowSelected[col].get("targetColumn")
                            + str(self.NewWorkBookCurrentRow)
                        ] = valueToPlace  # writes valueToPlace to New WB
                    if (self.skipRowIfMissingFlagg and not valueToPlace):
                        self.errors.append(f"Skipping row {i} {col}{self.StartRow + row} is missing given")
                        self.skippedRows -= 1 # decreasing the row counter so next row will not go passed the current
                        break
                    elif (
                        self.rowSelected[col].get("mandatory") and not valueToPlace
                    ):  # listing every field that is missing mandatory data
                        self.errors.append(
                            f"{col}{self.StartRow + row} is mandatory but no value given"
                        )
                        self.new_ws[
                            self.rowSelected[col].get("targetColumn")
                            + str(self.NewWorkBookCurrentRow)
                        ].fill = self.errorColor # Colors every field that is missing mandatory data

    def SaveFiles(self):
        if self.errors and len(self.errors) < 5:  # showing missing data as a messagebox warning
            messagebox.showwarning(
                title="Missing data in workbook",
                message=f"""{self.filename}\n{"\n".join(self.errors)}""",
            )
        elif self.errors:
            messagebox.showwarning(
                title="Missing data in workbook",
                message=f"""{self.filename}\nmissing data in {len(self.errors)} mandatory fields""",
            )
        saveFileName = filedialog.asksaveasfilename(
            title="Save file as",
            filetypes=[("Excel filer", "*.xlsx")],
            defaultextension=".xlsx",
        )
        if saveFileName:  # if no save file is set ignores save and open of Excel
            try:
                self.new_wb.save(saveFileName)
                savePath = Path(saveFileName)
                os.system(f'start EXCEL.EXE "{savePath}"')
            except PermissionError:
                messagebox.showwarning(title="Fil locked",message="Filen er alt åpen")
        self.new_wb.close()  # closes the new wb
        self.wb.close()  # closes the existing wb

    def __init__(self):
        self.errorColor = openpyxl.styles.fills.PatternFill(
            patternType="solid", fgColor=openpyxl.styles.colors.Color(rgb="00FF0000")
        )
        self.errors = []
        self.getInputData()
        self.variableSetup()
        self.skippedRows = 0
        if self.ws:
            self.setupNewFile()
            self.readAndWriteExcel()
            self.SaveFiles()
        else:
            messagebox.showwarning("Worksheet not found",message="Could not find the worksheet")

if __name__ == "__main__":
    transformExcel()
