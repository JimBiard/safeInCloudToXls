This simple python app creates an Excel workbook file from the contents of an exported SafeInCloud database XML file.
It does not track labels, nor does it process deleted or template entries.
It writes each Card and Note entry as a row in a single sheet in a .xlsx file.

The Title, Login, and Password fields are the first three columns in the sheet, the Notes field is the last column in the sheet, and all other fields found in the XML file follow the first three columns sorted in alphabetical order.

This python app depends on the openpyxl python package. It will work with Python 2 or Python 3.
