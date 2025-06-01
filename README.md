# excel-formatter
A python script that calculates subtotals of sections in an excel sheet and adds them to a summary sheet

features include;

- calculates total for every row 
- subtotals for sections that can be defined by you
- sending each subtotal to a summary sheet used to calculate a grand total

requirements;

- python 3.x

  download python from https://www.python.org/downloads/
  
- openpyxl

  get openpyxl by typing "pip install openpyxl" in cmd or powershell

how to use;

- edit the sections list in the script to meet the requirements of your spreadsheet
  
  the numbers in "start_row" and "end_row" are the row numbers in which each section starts and ends respectively (for example, if the items in "SECTION 1" start in row 1 and end in row 14, then "start_row" is 1 while "end_row" is 14)  "subtotal_row" is the cell in which you want the subtotal to display under the section and "summary_cell" being the cell in the summary sheet you want the subtotal of that section to be displayed in.

- edit the "qty_col" "price_col" and "total_col" to suit your spreadsheet's needs

  the numbers here are used to represent what columns your quantities, unit prices and total amount are in. a value of 1 is the A column, a value of 2 is the B column and so on, so adjust accordingly.

- edit wb.save to choose the preferred post formatting the name of your spreadsheet

  
