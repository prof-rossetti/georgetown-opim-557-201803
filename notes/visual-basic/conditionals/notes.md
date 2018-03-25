# Visual Basic Programming

## Conditionals

### If Statements

The [`If` Statement](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/ifthenelse-statement) allows the program to implement conditional logic.

All `If` statements contain an initial `If` clause and end with the keyword `End`. Everything inbetween is in the scope of that `If` statement.

```vb
If 5 = 5 Then 
  MsgBox("Hello") ' this statement will be executed, since the condition is true
End
```

```vb
If 5 = 4 Then 
  MsgBox("Hello") ' this statement will not be executed, since the condition is false
End
```

```vb
If 5 = 5 Then MsgBox("Hello") End ' a one-liner version, if you like that kind of thing
```

Add an `Else` clause to execute different code statements depending on whether or not a given condition is true:

```vb
If 5 = 4 Then
  MsgBox("Yep")
Else
  MsgBox("Nope")
End If
```

`If` statements can also contain any number of `Else If` clauses, each with their own condition. If there are multiple clauses that evaluate to true, the program will execute the first one:

```vb
If 5 = 4 Then
  MsgBox("Yep")
Else If 5 = 5 Then
  MsgBox("Other")
Else
  MsgBox("Nope")
End If
```
