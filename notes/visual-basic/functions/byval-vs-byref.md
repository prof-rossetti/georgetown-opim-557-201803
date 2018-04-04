# Visual Basic Programming

## Functions

### Definining Functions

#### `ByVal` Vs `ByRef`

Here is the difference, illustrated via a code example. Don't focus on the function return values - they will operate as expected. The part to focus on is how modifications to the passed parameter (`SomeMessage` in both cases) either do or don't persist after the respective functions are finished with execution.

```vb
Private Function ValMessage(ByVal SomeMessage As String)
    SomeMessage = "Val Val Val" ' <-- this variable modification doesn't persist after the function finishes execution
    ValMessage = "Some Return Value"
End Function

Private Function RefMessage(ByRef SomeMessage As String)
    SomeMessage = "Ref Ref Ref" ' <-- this variable modification persists even after the function finishes execution
    RefMessage = "Some Return Value"
End Function

Private Sub DoStuff()
    Dim OriginalMessage As String
    Dim OtherMessage As String
    Dim AnotherMessage As String
    
    OriginalMessage = "Original"
    MsgBox ("ORIGINAL: " & OriginalMessage) '--> "Original"
    
    OtherMessage = ValMessage(OriginalMessage)
    MsgBox ("ORIGINAL: " & OriginalMessage) '--> "Original"
    
    AnotherMessage = RefMessage(OriginalMessage)
    MsgBox ("ORIGINAL: " & OriginalMessage) '--> "Ref Ref Ref" <--- this is the difference when you pass a paramter by reference
End Sub
```