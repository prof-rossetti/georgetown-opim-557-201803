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

### Writing Values

To write a value to a cell:

```vb
Range("A1").Value = "fun times"
```

### Clearing Contents

To clear the contents of one or more cells:

```vb
Range("A1:C5").ClearContents
```

### Cells in a Range

Access all cells in a given range:

```vb
Range("A1:C5").Cells.Count ' --> 15
```

After studying loops, you can use one to iterate through all cells in a given range.