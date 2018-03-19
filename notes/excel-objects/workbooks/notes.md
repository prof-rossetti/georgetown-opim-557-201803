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

