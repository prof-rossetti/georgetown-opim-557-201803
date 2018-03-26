# Visual Basic Programming

## Datatypes

### Arrays

> "A set of sequentially indexed elements having the same intrinsic data type. Each element of an array has a unique identifying index number. Changes made to one element of an array don't affect the other elements." - [glossary of VBA terms](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/vbe-glossary)

An array represents an ordered list of zero or more items of the same datatype.

Individual array elements can be accessed by their position, or "index" value. Array indices are zero-based by default, meaning the index of the first element in an array is 0. However, arrays can be declared using different index values.

#### Declaration

[Declare a new array variable](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/declaring-arrays), optionally specifying the expected size and datatype of items it will contain:

```vb
Dim Teams(1 To 5) As String ' the array datatype is specified by the parentheses syntax, whereas the string datatype references the datatype of each item in the array.
```

Store items using an index number to indicate the item's unique position in the array:

```vb
Teams(1) = "New York Yankees"
Teams(2) = "New York Mets"
Teams(3) = "Boston Red Sox"
Teams(4) = "New Haven Ravens"
Teams(5) = "Washington Nationals"
```

Access an item by referencing its index number:

```vb
Teams(4) ' --> "New Haven Ravens"
```

Count items in an array by referencing the index of its first item (`LBound()`) and the index of its last item (`UBound()`):

```vb
UBound(MyList) - LBound(MyList) + 1 ' --> (5 - 1 + 1) --> 5 items
```

See also: [using arrays](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/using-arrays).

#### Iteration

Once you have studied loops, you can use a `For Each` loop to iterate through each item in an array:

```vba
Dim Team As Variant ' use Variant datatype instead of String datatype to avoid error. this line is only necessary if you have "Option Explicit" enabled.

For Each Team In Teams
  MsgBox (Team)
Next Team
```
