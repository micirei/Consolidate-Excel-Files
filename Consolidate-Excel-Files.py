#!/usr/bin/env python
import pathlib
import sys
import os, time
import pandas as pd
from rich.console import Console
console = Console()

def listFiles(format, exceptions=None):
    if format != "":
        return [file for file in os.listdir() if pathlib.Path(file).suffix == "." + format]
    else:
        return [os.listdir()]
    
def load_into_dataframe(file, targetDataFrame, targetSheet, startRow, keyColumn):
    print("")
    # Load DataFrame
    console.print(f"Loading into a DataFrame [yellow]{file}[/yellow]...")
    if targetSheet is not None:
        df = pd.read_excel(file, skiprows=startRow, sheet_name=targetSheet)
    else:
        df = pd.read_excel(file, skiprows=startRow)
    # Add column for file origin
    firstWordInFileName = (pathlib.Path(file).stem.split()[0]).upper()
    df.insert(0, 'Origine', firstWordInFileName)
    # Filter empty key column
    if keyColumn is not None:
        print("Filtering DataFrame...")
        df = df.dropna(subset=[df.columns[keyColumn]])
    print("DataFrame loaded, now combining with previous...")
    combinedDataframe = pd.concat([targetDataFrame, df], ignore_index=True)
    console.print(f"[green]Added {file} to the combined DataFrame.[/green]")
    return combinedDataframe

def dataFrameToExcel(dataFrame, file):
    print("")
    console.print("[yellow]Saving workbook...")
    dataFrame.to_excel(file, engine='xlsxwriter', index=False)

# INPUT
selectedFormat = str(input("File format (any): ")).strip().lower()
targetSheet = input("Sheet to import (active): ")
keyColumn = input("Column key (no key): ")
startRow = input("Starting row (1): ")
if not os.path.exists("OUTPUT"): 
    os.makedirs("OUTPUT")

# INPUT SANITAZING
if targetSheet == "": targetSheet = None
if keyColumn == "": keyColumn = None
if startRow == "": startRow = 1
try:
    if targetSheet is not None: targetSheet = str(targetSheet)
    if keyColumn is not None: keyColumn = int(keyColumn)
    if startRow is not None: startRow = int(startRow)
except ValueError:
    print("Invalid input.")
    exit = input("Press Enter to exit...")
    sys.exit()
startRow -= 1

# PROCESS
start_time = time.time()
outputFile = os.path.join("OUTPUT", "Output-" + selectedFormat.upper()+ ".xlsx")
filesToProcess = listFiles(selectedFormat)
combinedDataframe = pd.DataFrame()
for file in filesToProcess:
    combinedDataframe = load_into_dataframe(file, combinedDataframe, targetSheet, startRow, keyColumn)
dataFrameToExcel(combinedDataframe, outputFile)

# OUTPUT
console.rule("[green]FINISHED")
end_time = time.time()
elapsed_time = end_time - start_time
minutes = int((elapsed_time % 3600) // 60)
seconds = int(elapsed_time % 60)
print(f"Output file saved in {outputFile}.")
print("Elapsed time:", "{:02d}:{:02d}".format(minutes, seconds))

print("")
exit = input("Press Enter to exit...")