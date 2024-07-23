#!/usr/bin/env python
import pathlib
import sys
import os, time
import pandas as pd
from rich.console import Console
console = Console()

def listFiles(format, exceptions=None):
    if format:
        return [file for file in os.listdir() if pathlib.Path(file).suffix == "." + format]
    else:
        return [file for file in os.listdir() if file.endswith(('.xls', '.xlsx', '.xlsm', '.xlsb'))]
    
def load_into_dataframe(file, targetDataFrame, targetSheet, startRow, keyColumn, headers):
    print("")
    # Load DataFrame
    if file.endswith('.xlsx') or file.endswith('.xlsm'):
        engine = 'openpyxl'
    elif file.endswith('.xls'):
        engine = 'xlrd'
    elif file.endswith('.xlsb'):
        engine = 'pyxlsb'
    else:
        print(f"Unsupported file format")
        return None
    try:
        console.print(f"Loading into a DataFrame [yellow]{file}[/yellow]...")
        df = pd.read_excel(file, skiprows=startRow, sheet_name=targetSheet, engine=engine, header=headers)
    except ValueError as ve:
        print(f"ValueError: {ve}")
        return None
    except FileNotFoundError as fnfe:
        print(f"FileNotFoundError: {fnfe}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None
    # Add column for file origin
    firstWordInFileName = (pathlib.Path(file).stem.split()[0]).upper()
    try:
        df.insert(0, 'Origin', firstWordInFileName)
    except:
        console.print("[red]Couldn't add origin column to DataFrame.")
        return None
    # Filter empty key column
    if keyColumn is not None:
        print("Removing empty rows...")
        try:
            df = df.dropna(subset=[df.columns[keyColumn]])
        except:
            console.print(f"[red]Couldn't remove empty rows from column {keyColumn}.")
            return None
    try:
        print("DataFrame loaded, now combining with previous...")
        combinedDataframe = pd.concat([targetDataFrame, df], ignore_index=True)
    except:
        console.print(f"[red]Couldn't merge the two DataFrames. {keyColumn}")
        return None
    console.print(f"[green]Added {file} to the combined DataFrame.[/green]")
    return combinedDataframe

def dataFrameToExcel(dataFrame, file):
    print("")
    console.print("[yellow]Saving workbook...")
    dataFrame.to_excel(file, engine='xlsxwriter', index=False)

# INPUT
selectedFormat = str(input("File format (any): ")).strip().lower()
while selectedFormat not in ['xls', 'xlsx', 'xlsm', 'xlsb'] and selectedFormat != "":
    print(f"{selectedFormat} not supported. Use xls, xls, xlsm or xlsb.")
    selectedFormat = str(input("File format: ")).strip().lower()
targetSheet = input("Sheet to import (first): ")
keyColumn = input("Column key (no key): ")
startRow = input("Starting row (1): ")
headers = input("Headers (Yes): ")
if not os.path.exists("OUTPUT"): 
    os.makedirs("OUTPUT")

# INPUT SANITAZING
if selectedFormat == "": selectedFormat = None
if targetSheet == "": targetSheet = 0
if keyColumn == "": keyColumn = None
if startRow == "": startRow = 0
if str(headers).lower().strip() == "no":
    headers = None
else:
    headers = 0
try:
    if keyColumn is not None: keyColumn = int(keyColumn)
    if startRow is not None: startRow = int(startRow)
except ValueError:
    print("Invalid input.")
    exit = input("Press Enter to exit...")
    sys.exit()
startRow -= 1

# PROCESS
failedFiles = []
start_time = time.time()
filesToProcess = listFiles(selectedFormat)
combinedDataframe = pd.DataFrame()
for file in filesToProcess:
    newDataFrame = load_into_dataframe(file, combinedDataframe, targetSheet, startRow, keyColumn, headers)
    if newDataFrame is not None:
        combinedDataframe = newDataFrame
    else:
        failedFiles.append(file)
if selectedFormat:
    outputFile = os.path.join("OUTPUT", "Output-" + selectedFormat.upper()+ ".xlsx")
else:
    outputFile = os.path.join("OUTPUT", "OUTPUT-ANY.xlsx")
dataFrameToExcel(combinedDataframe, outputFile)

# OUTPUT
console.rule("[green]FINISHED")
end_time = time.time()
elapsed_time = end_time - start_time
minutes = int((elapsed_time % 3600) // 60)
seconds = int(elapsed_time % 60)
print(f"Output file saved in {outputFile}.")
print("Elapsed time:", "{:02d}:{:02d}".format(minutes, seconds))
if failedFiles: console.print(f"Failed files: {list(failedFiles)}")

print("")
exit = input("Press Enter to exit...")