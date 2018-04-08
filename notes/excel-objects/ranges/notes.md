# MS Excel Objects

## Ranges

The [`Range`](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/range-object-excel) object represents one or more cells in a given worksheet.

### Reading Values and Properties

To read the value and other properties of a cell:

```vb
Range("A1").Value ' --> "Hello World"
Range("A1").Address ' --> "$A$1"
Range("A2").Formula ' --> "=B2+C2"
```

By default, ranges are referenced relative to the current sheet. If you need to reference a range on another sheet or a specific sheet, include the sheet name as part of the reference:

```vb
Worksheets("Sheet1").Range("A1").Value ' --> "Hello from Sheet 1"
Worksheets("Sheet2").Range("A1").Value ' --> "Hello from Sheet 2"
```

To detect the range of used cells on any given sheet, reference the `UsedRange` property:

```vb
Worksheets("Sheet1").UsedRange.Address ' --> $A$1:$L$36
```

> Warning: if a cell looks empty (i.e. it has no contents) but contains formatting, it will still be included in the `UsedRange`.

### Writing Values

To write a value to a cell:

```vb
Range("A1").Value = "fun times"
```

### Clearing Contents

To clear the contents of one or more cells:

```vb
Range("A1:C5").ClearContents ' clears contents, but does not clear formatting
Range("A1:C5").Clear ' clears all contents and formatting
```

### Cells in a Range

Access all cells in a given range:

```vb
Range("A1:C5").Cells.Count ' --> 15
```

After studying loops, you can use one to iterate through all cells in a given range:

```vb
Dim MyCell As Range

For Each MyCell In Range("A1:C5").Cells
  MsgBox (MyCell.Address)
Next MyCell
```

### Copying Ranges

To copy the contents of one range of cells to another, simultaneously read and write to and from the appropriate ranges:

```vb
Range("A1").Value = Range("B1").Value ' copies contents of B1 into A1
```

You can do this for multiple cells, or even entire rows/columns:

```vb
Range("A1:A10").Value = Range("B1:B10").Value ' copies contents of B1:B10 into range A1:A10
Range("A:A").Value = Range("B:B").Value ' copies contents of column B into column A
```

You can also do this from one workbook or worksheet to another:

```vb
Worksheets("Sheet1").Range("A1").Value = Worksheets("Sheet2").Range("A1").Value ' copies contents of A1 on Sheet2 into A1 on Sheet1
```
