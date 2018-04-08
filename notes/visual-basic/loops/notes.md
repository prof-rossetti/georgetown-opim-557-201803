# Visual Basic Programming

## Loops

Reference:

  + [Loops](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/looping-through-code)
  + [`Do ... Loop` loops](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/using-doloop-statements)
  + [`For ... Next` loops](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/using-fornext-statements)
  + [`For Each ... Next` loops](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/using-for-eachnext-statements)

In computer programs, loops can be used to iteratively execute one or more statements of code. Loops can help iterate through a collection of items, processing one item at a time. Loops can also help perform a task a certain number of times in succession.

In VBA, there are four kinds of loops:

  + `Do While ... Loop`: Repeat a statement as long as a condition is true.
  + `Do Until ... Loop`: Repeat a statement until a condition is true.
  + `For ... Next`: Repeat a statement a specified number of times.
  + `For Each ... Next`: Repeat a statement for each object in a collection of objects.

To programmatically exit from a `Do` loop, use the statement `Exit Do`. To programmatically exit from a `For` loop, use the statement `Exit For`. To manually exit a loop, for example if you get stuck in an infinite loop, press and hold the "Escape" key to exit the program. If that doesn't work, try pressing the "Ctrl" + "Break" keys. If that doesn't work, try "Ctrl" + "Alt" + "Delete" to reveal the task manager and force quit the application.

### `Do While ... Loop` Loops

This kind of loop will continue **as long as** a certain logical condition is met. In other words, it will stop when the condition is no longer being met.

```vb
Dim Counter As Integer
Counter = 1

Do While Counter <= 5
  MsgBox("The counter's value is currently: " & Counter)
  Counter = Counter + 1 ' increment to avoid an infinite loop!!!!
Loop
```

### `Do Until ... Loop` Loops

This kind of loop will continue **until** a certain logical condition is met. In other words, it will stop when the condition is met.

```vb
Dim Counter As Integer
Counter = 1

Do Until Counter = 5
  MsgBox("The counter's value is currently: " & Counter)
  Counter = Counter + 1 ' increment to avoid an infinite loop!!!!
Loop
```

### `For ... Next` Loops

This kind of loop will repeat a statement a specific amount of times. The counter incrementing mechanism is built-in to the loop's syntax.

```vb
Dim Counter As Integer

For Counter = 1 To 5 ' specify the number of times this loop will repeat
  MsgBox("The counter's value is currently: " & Counter)
Next Counter ' increment the Counter's value and execute the next iteration
```

### `For Each ... Next` Loops

This kind of loop will iterate over each object in a collection of objects. Examples of collections include a [range](/notes/excel-objects/ranges/notes.md) of cell objects, as well as an [array](/notes/visual-basic/datatypes/arrays.md) of items.

```vb
Dim MyCell As Range
Dim MyRange As Range

Set MyRange = Range("A1:C5")

For Each MyCell in MyRange.Cells
  MyCell.Value = MyCell.Address ' an example of something to do with MyCell
Next
```

See also: [looping through items in an array](/notes/visual-basic/datatypes/arrays.md#iteration).
