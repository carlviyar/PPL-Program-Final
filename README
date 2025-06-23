# Overview

This program presents a GUI that enables the user to "check off" labs for a patient's lab order like they would if they were manually checking off the labs on paper. The program then takes the specified labs, pulls lab data and prices from an Excel file, calculates a total price for all the labs, and then outputs the results into another Excel file that serves as an invoice for the labs.

This program streamlines the lab calculation and invoice creation process to save time and minimize human error while also ensuring that the lab codes/names/prices are easily editable by a non-programmer (via the Excel sheets).

# Tech specifications

### Key files:

- main.py contains actual code for tkinter (design) + processing data
- excel.py is a module that contains code for grabbing tests and prices
  from "PPL Prices Reference.xlsx" as well as outputting the final PPL invoice excel file at the end

### Required python modules:

- customtkinter
- tkinter
- datetime
- os
- openpyxl
- pyinstaller

## How to run

To run without creating executable, run:\
`python cli.py`

### Outputting .exe file:

To export one singular .exe file, we are using the "pyinstaller" module.\
`pyinstaller cli.py --name PPL_Program --onefile --noconsole`

The exported file can be found under `./dist/PPL_Program.exe`\
**You must MOVE this executable file to the root directory (to properly grab from the excel data sheets)**

`carl.viyar@gmail.com` is the owner of the repo. Any questions may be directed to him.
