# SeleniumExcelApp
Scrapes financial adviser's lead info off of a lead website and updates an Excel spreadsheet with data.

I plan to create a Docker file for it.


### To properly set up Excel workbook to interface with this app:

1. Cell B1 of the first (leftmost) spreadsheet is used as the "Last Updated" cell. It tells the app when the app last successfully ran.
2. Cell D1 of the first (leftmost) spreadsheet is used as the "Run Errors" cell. It tells the user if there was an error at various points in the process. For instance, the app can tell when there is a webdriver error, login error, spreadsheet format error, cell format error, etc.
3. Cell F1 for each respective spreadsheet is used as the "Name Error" cell. It lists the names that were on the webpage but not in the respective spreadsheet for a given Order Number.
4. Each spreadsheet should be named after the correspoding Order Number.
5. Each spreadsheet should have a table with all of the client data, and that table should be named: 'Table'+str(CorrespondingOrderNumber)
6. Each table should have the folloing columns:
  - Column B is for Received dates. This is is a datetime of the form: '%Y-%m-%d %H:%M%p'
  - Column D is First Name
  - Column E is Last Name
  - Column J is Zip Code
  - Column M is JPG Hyperlink
