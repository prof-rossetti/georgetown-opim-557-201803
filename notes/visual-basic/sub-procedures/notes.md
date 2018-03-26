# Visual Basic Programming

## Sub-procedures

In Visual Basic, a [Sub-procedure](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/procedures/sub-procedures) defines a subset of application logic that will be executed when the procedure is invoked.

Whereas the responsibility of a [function](/notes/visual-basic/functions/notes.md) is to return some value, the responsibility of a sub-procedure is to simply "perform a task".

### Defining Sub-procedures

```vb
Private Sub DoStuff()
  ' write code here which will perform some action
End Sub
```

```vb
Private Sub DisplayMyMessage()
  MsgBox("My message is: Hello World")
End Sub
```

Sub-procedure definitions begin with the statement `Private Sub`, followed on the same line by the name of the sub-procedure (in this case `DisplayMyMessage()`), followed by one or more lines of indented code, and finally concluding with the statement `End Sub`.

To programmatically exit from a sub-procedure, use the statement `Exit Sub`.

Note the trailing parentheses in the sub-procedure's name. They not only visually indicate this statement is a procedure, but they also serve as a space to pass parameters (see below).

#### Defining Sub-procedures with Parameters

When necessary and appropriate, specify one or more arguments (a.k.a. "parameters"), inside the parentheses part of the sub-procedure definition. The syntax for defining parameters is similar to the syntax for declaring variables, except a different keyword is used (either `ByVal` or `ByRef`). Use `ByVal` in most cases, but use `ByRef` if you need changes to the parameter to remain in memory after the function has finished execution.

```vb
Private Sub DisplayCustomMessage(ByVal SomeMessage As String)
  MsgBox("The custom message is: " & SomeMessage)
End Sub
```

These defined parameters represent variable values that are expected to be passed to the sub-procedure when it is invoked.

### Invoking Sub-procedures

The code inside a sub-procedure won't execute until/unless invoked. Sub-procedures are generally invoked when the user triggers an event (like a button click event), or when the user "runs" the program.

However, you can also invoke a sub-procedure programmatically by using the `Call` keyword, followed by the name of the sub-procedure:

```vb
Call DisplayMyMessage ' --> a message box pops up with the content... "My message is: Hello World"
```

```vb
Call DisplayCustomMessage("Hello World") ' --> a message box pops up with the content... "The custom message is: Hello World"

Call DisplayCustomMessage("Goodbye") ' --> a message box pops up with the content... "The custom message is: Goodbye"
```
