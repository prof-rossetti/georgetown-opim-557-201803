# Issuing Web Requests

## CSV Challenges

### CSV Challenge 1: Teams

Write a VBA program which issues a GET request for this example CSV-formatted [teams data](https://raw.githubusercontent.com/prof-rossetti/georgetown-opim-557-201803/master/exercises/web-requests/data/teams.csv), then write the results to a corresponding range of spreadsheet cells.

### CSV Challenge 2: Gradebook

Write a VBA program which issues a GET request for this example CSV-formatted [gradebook data](https://raw.githubusercontent.com/prof-rossetti/georgetown-opim-557-201803/master/exercises/web-requests/data/gradebook.csv), then write the results to a corresponding range of spreadsheet cells, then calculate and display in a message box the average, min, and max grades.

<hr>

## JSON Challenges

Unless you already have a preferred way of parsing JSON in VBA, let's try this open source module called [VBA-JSON](https://github.com/VBA-tools/VBA-JSON).  Installation instructions are in that repository's documentation. :octocat:

After issuing an HTTP request, if your response text looks like JSON, try parsing it using `JsonConverter.ParseJson()`:

```vb
Dim ResponseObj As Object
Set ResponseObj = JsonConverter.ParseJson(MyResponseText)
MsgBox (TypeName(ResponseObj)) '--> ??? Dictionary or Collection, etc.
```

You will have to process the top-level response data differently depending on whether it represents an object (`Dictionary`) or an array of objects (`Collection`).

> Hint: Collections can be looped through, and here's a [reference document on Dictionaries](/notes/visual-basic/datatypes/dictionaries.md).

### JSON Challenge 1: Team

Write a VBA program which issues a GET request for this example JSON-formatted [team data](https://raw.githubusercontent.com/prof-rossetti/georgetown-opim-557-201803/master/exercises/web-requests/data/teams/1.json), then write the results to a corresponding range of spreadsheet cells.

> Hint: the response is a JSON Object.

### JSON Challenge 2: Teams

Write a VBA program which issues a GET request for this example JSON-formatted [teams data](https://raw.githubusercontent.com/prof-rossetti/georgetown-opim-557-201803/master/exercises/web-requests/data/teams.json), then write the results to a corresponding range of spreadsheet cells.

> Hint: the response is a JSON Array of Objects. :smile_cat:

### JSON Challenge 3: Gradebook

Write a VBA program which issues a GET request for this example JSON-formatted [gradebook data](https://raw.githubusercontent.com/prof-rossetti/georgetown-opim-557-201803/master/exercises/web-requests/data/gradebook.json), then write the results to a corresponding range of spreadsheet cells, then calculate and display in a message box the average, min, and max grades.

> Hint: the response is a JSON Object with a nested Array. :smile_cat: :smile_cat:
