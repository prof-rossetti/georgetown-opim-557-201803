# MS Excel Objects

## Ranges

The [`Range`](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/range-object-excel) object represents one or more cells in a given worksheet.

### Reading Values and Properties

To read the value of a cell:

```vb
Range("A1").Value ' --> "Hello World"
```

To read various properties of a cell:

```vb
Range("A1").Address ' --> "$A$1"
Range("A2").Formula ' --> "=B2+C2"
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