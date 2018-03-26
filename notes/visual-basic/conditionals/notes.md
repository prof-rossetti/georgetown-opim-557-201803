# Visual Basic Programming

## Conditionals

### If Statements

The [`If` Statement](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/ifthenelse-statement) allows the program to implement conditional logic.

All `If` statements contain an initial `If` clause and generally end with the keywords `End If`. Everything inbetween is considered to be inside the scope of that `If` statement:

```vb
If 5 = 5 Then 
  MsgBox("Hello") ' this statement will be executed, since the condition is true
End If
```

```vb
If 5 = 4 Then 
  MsgBox("Hello") ' this statement will not be executed, since the condition is false
End If
```

```vb
If 5 = 5 Then MsgBox("Hello") ' a one-liner version, if you like that kind of thing
```

Add a final `Else` clause to execute different code statements depending on whether or not a given condition is true:

```vb
If 5 = 4 Then
  MsgBox("Yep")
Else
  MsgBox("Nope")
End If
```

Add any number of `ElseIf` clauses, each with their own condition. If there are multiple clauses that evaluate to true, the program will execute the first one:

```vb
If 5 = 4 Then
  MsgBox("Yep")
ElseIf 5 = 5 Then
  MsgBox("First true condition") ' this statement will always get executed
ElseIf True = True Then
  MsgBox("Second true condition") ' this statement will never get executed
Else
  MsgBox("Nope")
End If
```
