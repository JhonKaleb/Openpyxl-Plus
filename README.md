<h1 align="center">
  Openpyxl Plus
  <br>
</h1>

<h4 align="center">New functionalities and facilities for the <a href="https://openpyxl.readthedocs.io/en/stable/tutorial.html target="_blank">Openpyxl Library</a>.</h4>

<p align="center">
  <a href="#key-features">Key Features</a> •
  <a href="#how-to-use">How To Use</a> •
  <a href="#credits">Credits</a>
</p>

## Key Features

* Set values - You can set values in any columns and rows in a easy way
* Create headers - You can easily insert an header in any sheet or all sheets at once.
* Merge multiple rows- You can merge multiple rows in an range at once.

## How To Use

To clone and run this library extension, you'll need [Git](https://git-scm.com) and [Openpyxl](https://pypi.org/project/openpyxl/) (which comes with pip in python) installed on your environment. From your command line:

### Cloning and installing

```bash
# Clone this repository
$ git clone https://github.com/JhonKaleb/Openpyxl-Plus

# Go into the repository folder
$ cd Openpyxl-Plus

# Install dependencies
$ pip install openpyxl

# Or using requirements.txt
$ pip install -r requirements.txt

# If you already have openpyxl in your project, you can copy the files in this repository folder to your openpyxl library folder:
# Move workbook_plus.py to site-packages/openpyxl/workbook
# Move worksheet_plus.py to site-packages/openpyxl/worksheet
```

### Usage exemple
#### Simple header and table
Below is an example of how to generate a table with a simple header from an array.

``` Python
# Importing the workbookplus
from openpyxl.workbook.workbook_plus import WorkbookPlus

# Creating an workbook and an worksheet
workbook = WorkbookPlus()

# When you start a WorkbookPlus object, its already comes with an sheet
worksheet = workbook.active

# But if you need more sheets, you can do this:
# worksheet = workbook.create_sheet("sheet name")

header_full_row = [["Fruits price and quantity"]]
header_table = [["Fruit", "Price", "Quantity"]]

fruits = [
    ["apple", 10.99, 50],
    ["banana", 15.99, 30],
    ["Strawberry", 20.99, 40]
    ]

# Setting fruits table in the worksheet
worksheet.set_values_with_array(fruits)

# Setting the two headers with diferent stylization
worksheet.set_header(header_table, merge_header_entire_row=False)
worksheet.set_header(header_full_row, merge_header_entire_row=True)

# Saving the spreadsheet generated with Fruits_sheet.xlsx file name
workbook.save("Fruits_sheet.xlsx")

```
This source code will generate the sheet below.

![usage-1-result](https://github.com/JhonKaleb/utils-repository/blob/main/openpyxl-plus/usage-1-result.png)

#### Seting values to a specific column and persist header

Exemple of to how merge multiple rows and create a header that persists on all tabs of the spreadsheet.
``` Python
# Importing the workbookplus
from openpyxl.workbook.workbook_plus import WorkbookPlus

workbook = WorkbookPlus()
worksheet_1 = workbook.active
worksheet_2 = workbook.create_sheet("second sheet")

header = [["Sales in 2019"]]
column1 = ["jan", "feb", "mar"]
column2 = [1000, 2000, 15000]

# Setting values to an specifc column
worksheet_1.set_column_values(column1, column=1)
worksheet_1.set_column_values(column1, column=2)

# Setting a header to both sheet_1 and sheet_2
workbook.set_header_in_all_sheets(header, merge_header_entire_row=True)

# Saving the spreadsheet generated with Sales.xlsx file name
workbook.save("Sales.xlsx")
```
This source code will generate the sheet below.

* Tabs

![usage-2-result-tabs](https://github.com/JhonKaleb/utils-repository/blob/main/openpyxl-plus/usage-2-result-tabs.png)

* Default sheet

![usage-2-result](https://github.com/JhonKaleb/utils-repository/blob/main/openpyxl-plus/usage-2-result.png)

* Second Sheet

![usage-2-results-second-sheet](https://github.com/JhonKaleb/utils-repository/blob/main/openpyxl-plus/usage-2-results-second-sheet.png)


#### Seting values to a specifc row and Merging ranges in multiple rows
``` Python
# Importing the workbookplus
from openpyxl.workbook.workbook_plus import WorkbookPlus

workbook = WorkbookPlus()
worksheet_1 = workbook.active

row1 = ["jan", "feb", "/mar/aph/mai"]
row2 = [1000, 12000, 4000]
row3 = [2000, 13000, 5000]
row4 = [3000, 14000, 6000]


# Setting values to an specifc row
worksheet_1.set_row_values(row1, row=1)
worksheet_1.set_row_values(row2, row=2)
worksheet_1.set_row_values(row3, row=3)
worksheet_1.set_row_values(row4, row=4)
# In this case is better use set_values_with_array(), but is just an method example :D

# Merging each row 1 to 4, from te column 3(C) to column 5(E)
worksheet_1.merge_range(3, 5, 1, 4)

workbook.save("test.xlsx")
```
This source code will generate the sheet below.

![usage-3-result](https://github.com/JhonKaleb/utils-repository/blob/main/openpyxl-plus/usage-3-result.png)


## Emailware

Openpyxl-plus is an [emailware](https://en.wiktionary.org/wiki/emailware). Meaning, if you liked using this project or it has helped you in any way, I'd like you send me an email at <jhon.kaleb@hotmail.com> about anything you'd want to say about this extention. I'd really appreciate it!

## Credits

This software uses the following open source packages:

- [Openpyxl](https://openpyxl.readthedocs.io/en/stable/)
