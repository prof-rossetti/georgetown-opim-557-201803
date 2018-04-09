# Processing CSV Strings

## Challenge

Write VBA code that will process the following Comma-Separated Values (CSV) string into a corresponding spreadsheet of cells.

## Instructions

Open a new workbook, rename the first sheet to "Interface, insert on it a command button on the first sheet, and inside its click event sub-procedure, paste the code below:

```vb
Dim MyStr As String

MyStr = "city,name,league" & vbNewLine & _
        "New York,Mets,Major" & vbNewLine & _
        "New York,Yankees,Major" & vbNewLine & _
        "Boston,Red Sox,Major" & vbNewLine & _
        "Washington,Nationals,Major" & vbNewLine & _
        "New Haven,Ravens,Minor"

MsgBox(MyStr)

' TODO: write some VBA code here!
```

Create another sheet called "Data".

Write code inside the command button's click event sub-procedure that will clear the contents of the "Data" sheet, write the desired spreadsheet output there (see below), and finally activate that sheet:

city | name | league
--- | --- | ---
New York | Mets | Major
New York | Yankees | Major
Boston | Red Sox | Major
Washington | Nationals | Major
New Haven | Ravens | Minor
