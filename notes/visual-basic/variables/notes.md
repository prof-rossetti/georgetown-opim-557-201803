# Visual Basic Programming

## Variables

### Declaring Variables

Visual Basic has traditionally been a "statically-typed" language, which means it requires the developer to declare which [type of data](/notes/visual-basic/datatypes/notes.md) a variable will hold. In the current version it is not always necessary to declare variables to produce desired functionality. However, declaring variables is a best practice, at least for performance reasons.

The most common way to [declare a variable](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/dim-statement) is to use the `Dim` keyword.
, followed by the variable name, followed by the `As` keyword, followed by the datatype. For example:

```vb
Dim MyNumber As Integer

Dim MyText As String

Dim MyDecimal As Double

Dim MyBool As Boolean

Dim MyDate As Date

Dim MySheet As Worksheet

Dim MyCell As Range
```

### Assigning Values to Variables

Use an equality operator (`=`) to assign some value on the right side of the `=` to a given variable on the left side of the `=`. For example:

```vb
MyNumber = 25

MyText = "Hello World"

MyDecimal = 3.14

MyDate = #10/31/2017# ' the pound signs surround the date value formatted as MM/DD/YYYY
```

To assign an [Excel Object](/notes/excel-objects) to a variable, you may need to use the `Set` keyword:

```vb
Set MySheet = Worksheets("Sheet1")

Set MyCell = Range("C5")
```

### Referencing Variables

After a variable is declared and assigned, any subsequent references to the variable name will yield the variable's value:

```vb
"All I have to say is: " & MyText ` --> "All I have to say is: Hello World"

MyNumber + MyDecimal ` --> 28.14
```

And when applicable, references to the variable will also provide access to its properties:

```vb
MySheet.Name ' --> "Sheet1"

MyCell.Address ' --> "$C$5"
```
