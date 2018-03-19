# MS Excel Objects

## Worksheets

The [`Worksheet`](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/worksheet-object-excel) object references a corresponding worksheet in the workbook.

To access a specific sheet, reference it by name (e.g. "Sheet1") or position in the workbook (e.g. 1):

```vb
Worksheets("Sheet1").Name ' --> "Sheet1"
Worksheets("Sheet1").Index ' --> 1
```

Like the `Range` object, the `Worksheet` object also has a `Cells` property, which can be used to manipulate the sheet's cell values:

```vb
Worksheets("Sheet1").Cells.ClearContents ' remove values of all cells in this sheet
```

Pass a row number and a column number to reference a specific cell:

```vb
Worksheets("Sheet1").Cells(1,3).Value = "good stuff" ' where 1 is the row number and 3 is the column number (a.k.a. cell "C1")
```

#### Creating Worksheets

To [create](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/sheets-add-method-excel) a new sheet:

```vb
Worksheets.Add
```

#### Activating Worksheets

To switch a user's active view to a given worksheet: 

```vb
Worksheets("Sheet3").Activate
```

#### Looping Through Worksheets

After you have studied looping and arrays, you can apply the concepts to loop through a collection of worksheets:

```vb
' adapted from source: https://www.extendoffice.com/documents/excel/3057-excel-delete-all-sheets-except-current.html

Dim MySheet As Worksheet

For Each MySheet In Worksheets
    MsgBox (MySheet.Name)
Next
```

#### Deleting Worksheets

To [delete](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/worksheet-delete-method-excel) a given sheet:

```vb
Worksheets("Sheet3").Delete
```
