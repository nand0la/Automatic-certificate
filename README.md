# certificate automation
Python script that assembles certificates based on an excel spreadsheet

## modules
- PIL: for image editing
- xlrd: to edit spreadsheets

<hr>

## How to use
- Open to files
- Put the templates in the "template_certificate" folder
- Coloque a planilha na pasta "planilha_alunos"

Depending on the sheet used in the spreadsheet, the 
```python
    def return_list_students(file):
        table_stud = x.open_workbook(f'planilha_alunos/{arquivo}')
        sheet_students = table_stud.sheet_by_index(0)  # <- change the number of the sheet_by_index method (x)
        num_rows = sheet_students.nrows
        ...
```
