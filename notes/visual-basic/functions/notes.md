# Visual Basic Programming

## Functions

In Visual Basic, a [Function](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/procedures/function-procedures) is a specific kind of procedure which is responsible for returning some value.

### Defining Functions

```vb
Private Function MyMessage()
  MyMessage = "My message is: Hello World"
End Function
```

Function definitions begin with the statement `Private Function`, followed on the same line by the name of the function (in this case `MyMessage()`), followed by one or more lines of indented code, and finally concluding with the statement `End Function`.

Note the final, "return", variable name needs to be the same as the function name (e.g. `MyMessage`).

Also note the trailing parentheses in the function's name. They not only visually indicate this statement is a procedure, but they also serve as a space to pass "parameters" (see section below).

#### Defining Functions with Parameters

When necessary and appropriate, specify one or more arguments (a.k.a. "parameters"), inside the parentheses part of the function definition. The syntax for defining parameters is similar to the syntax for declaring variables, except a different keyword is used (either `ByVal` or `ByRef`). Use `ByVal` in most cases, but use `ByRef` if you need changes to the parameter to remain in memory after the function has finished execution. If you're curious, after you finish this document, take a look at an example illustrating the [differences between `ByVal` and `ByRef`](/byval-vs-byref.md).

```vb
Private Function CustomMessage(ByVal SomeMessage As String)
  CustomMessage = "The custom message is: " & SomeMessage
End Function
```

```vb
Private Function RectangleArea(ByVal Length As Integer, ByVal Width As Integer)
  RectangleArea = Length * Width
End Function
```

```vb
Private Function FullName(ByVal FirstName As String, ByVal LastName As String)
  FullName = FirstName & " " & LastName
End Function
```

These defined parameters represent variable values that are expected to be passed to the function when it is invoked.

### Invoking Functions

The code inside a function won't execute until/unless invoked. Invoke a function by referencing its name.

```vb
MyMessage() ' --> "My message is: Hello World"
```

```vb
CustomMessage("Hello World") ' --> "The custom message is: Hello World"

CustomMessage("Goodbye") ' --> "The custom message is: Goodbye"
```

```vb
RectangeArea(10, 7) ' --> 70

RectangeArea(2, 3) ' --> 6
```

```vb
FullName("Ryan", "Chan") ' --> "Ryan Chan"

FullName("Chan", "Ryan") ' --> "Chan Ryan"
```

As you may have noticed, parameter order is important when invoking functions. But if you'd like more flexibility, you can invoke a function using named parameters in whatever order you prefer:

```vb
FullName(LastName:="Chan", FirstName:="Ryan") ' --> "Ryan Chan"
```
