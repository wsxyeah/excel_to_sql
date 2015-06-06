# excel_to_sql
To gernerate SQL 'insert' statements from excel data

## Dependencies
    pip install xlrd

## Excel file requirements
- Support .xls and .xlsx file
- Every sheet should correspond to one table in your database
- The first row of one sheet should be the colum names of the table
- You should set the excel data type to 'TEXT'