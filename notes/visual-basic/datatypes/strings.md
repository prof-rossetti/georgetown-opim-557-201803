# Visual Basic Programming

## Datatypes

### Strings

The [String](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/string-data-type) datatype is used to represent words or text. A string must begin with an opening quotation mark (`"`) and end with a closing quotation mark (`"`) (e.g. `"Hello World"`).

```vb
Dim MyMessage As String
MyMessage = "Hello World"
MsgBox(MyMessage)
```

#### String Operations

Just like you can perform arithmetic operations on numbers, you can perform designated [string operations](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/functions/string-functions)
 on strings.

##### String Concatenation

The most popular string operation is "concatenation", which assembles multiple strings into a single string. The operator to perform string concatenation is an ampersand (`&`). When concatenating strings with other strings, or even with variables, make sure to include space characters in the proper places or else your strings will run together without a space. For example, **all the following approaches are equivalent**:

```vb
Dim MyMessage As String
MyMessage = "Hello" & " " & "World" ' notice the separate space character
MsgBox(MyMessage)
```

```vb
Dim MyMessage As String
MyMessage = "Hello " & "World" ' notice the trailing space after the word Hello
MsgBox(MyMessage)
```

```vb
Dim MyMessage As String
MyMessage = "Hello" & " World" ' notice the leading space before the word World
MsgBox(MyMessage)
```

```vb
Dim FirstString As String
Dim SecondString As String
MyMessage = FirstString & " " & SecondString ' notice the separate space character in-between the two variables. just because you use variables to represent strings does not change your need to include space characters
MsgBox(MyMessage)
```

##### New Lines

Use the `vbNewLine` keyword to insert a line break in a concatenated string:

```vb
Dim MyMessage As String
MyMessage = "Hello World" & vbNewLine & "Goodbye!"
MsgBox(MyMessage)
```

> Note: a new line character can be represented in VBA by any of the following: `vbLf`, `vbCrLf`, `vbCr`, and `vbNewLine`.

##### String Case

Use the `UCase()`, `LCase()` and `WorksheetFunction.Proper()` functions to manipulate the case of any string:

```vb
Dim MyMessage As String
MyMessage = "HeLlo WoRlD"
MsgBox(MyMessage & " " & UCase(MyMessage) & " " & LCase(MyMessage) & " " & WorksheetFunction.Proper(MyMessage))
```

##### String Formatting

Use the built-in [`Format()` function](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/format-function-visual-basic-for-applications) to convert numbers into strings using a specified template. Two common templates are `"Currency"` and `"Percent"`, however you can also create your own custom formats using a mix of special characters:

```vb
Dim Price As Double
Price = 45.12345

Format(Price, "Currency") '--> "$45.12"
Format(Price, "Percent") '--> "4512.34%"
Format(Price, "##,##0.0 tons") '--> "45.1 tons"
```

##### Substring Detection

Use the `InStr()` function to detect whether or not a string includes a specified substring. The first parameter represents the string to be searched, and the second parameter represents the substring to search for.

If the substring is found, the function will return a `0`, otherwise it will return the substring index number representing the position of the substring's first character:

```vb
Dim MyStr As String
MyStr = "Hello World"

InStr(MyStr, "World") ' --> 7

InStr(MyStr, "Goodbye") ' --> 0
```

<hr>

> Below this line there are advanced topics which you can feel free to come back to later, especially once you have studied arrays...

<hr>

##### String Splitting

Split a string into component parts by using the `Split()` function and passing parameters corresponding to the string to be split, followed by the delimiter:

```vb
Dim MyStr As String
MyStr = "first | second | third"

Dim MyList() As String ' an array of strings
MyList = Split(MyStr, " | ")
```

When a string is split, it results in an [array](arrays.md), which can be accessed in the usual ways:

```vb
Dim ListItem As Variant

For Each ListItem In MyList
    MsgBox(ListItem)
Next ListItem

' --> "first"
' --> "second"
' --> "third"
```
