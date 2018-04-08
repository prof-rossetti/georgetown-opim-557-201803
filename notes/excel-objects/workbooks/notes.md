# MS Excel Objects

## Workbooks

The [`Workbook`](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/workbook-object-excel) object represents a file containing one or more worksheets.

To access a specific workbook, reference its name (e.g. "Sheet1") or its position among open workbooks (e.g. 1). Or reference the `ActiveWorkbook`:

```vb
Workbooks("mybook.xlsm").Name ' --> "mybook.xlsm"
Workbooks(1).Name ' --> "mybook.xlsm"
ActiveWorkbook.Name ' --> "mybook.xlsm"
```

To activate a given workbook:

```vb
Workbooks("mybook.xlsm").Activate
```

### Opening and Closing Workbooks

To open a workbook, try using the in [`Application.Workbooks.Open()` method](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/workbooks-open-method-excel):

```vb
Dim MyWorkbook As Workbook
Set MyWorkbook = Application.Workbooks.Open(SomeFileName) ' where SomeFileName is the path of a local file openable by MS Excel
```

> Note: pay careful attention to which workbook is considered "active" when dealing with multiple workbooks at the same time.

To close a workbook:

```vb
MyWorkbook.Close
```

#### Selecting Files to Open

When opening workbooks and CSV files, you might want to use the [`Application.GetOpenFilename()` method](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/application-getopenfilename-method-excel) to allow the user to select an existing file and ensure a proper file name gets passed to `Application.Workbooks.Open()`:

```vb
Dim SelectedFileName As String
SelectedFileName = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Please select a CSV file representing monthly sales data...")

Dim MyWorkbook As Workbook
Set MyWorkbook = Application.Workbooks.Open(SelectedFileName)
```

> Note: This method accepts a "file filter" as its first parameter and a dialogue box title as its third parameter. Two common file filters you may need to use include: `"Text Files (*.csv),*.csv"` and `"Text files (*.xlsx),*.xlsx"`.

> Note: If the user presses "Cancel" instead of selecting a file, the resulting return value will be `False` or `"False"` instead of the file name.
