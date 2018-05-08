# Visual Basic Programming

## Datatypes

### Dictionaries

#### Setup

To use the `Dictionary` datatype in VBA, you may need to enable use of the Microsoft Scripting Runtime from the VBE menu:
"Tools" > "References" > "Microsoft Scripting Runtime".

#### Initialization

Create a new dictionary and set its attributes via the `Add` method.

```vb
Dim MyObj As Object
Set MyObj = CreateObject("Scripting.Dictionary")
MsgBox (TypeName(MyObj)) ' --> Dictionary

MyObj.Add "day", "Tuesday"
MyObj.Add "time", "Morning"
```

The `Add` method takes two parameters: the first being the name of the attribute (also known as the "key" - e.g. "day"), and the second is the actual attribute "value" (e.g. "Tuesday").

#### Accessing Attributes

Access any attribute by passing its key as a string parameter to the dictionary itself:

```vb
MsgBox (MyObj("day")) ' --> Tuesday
MsgBox (MyObj("time")) ' --> Morning
```

#### Iteration

It is possible to iterate over a dictionary's keys, which are themselves a collection:

```vb
Dim MyKey
For Each MyKey In MyObj
    MsgBox (MyKey & ": " & MyObj(MyKey))
Next MyKey
```
