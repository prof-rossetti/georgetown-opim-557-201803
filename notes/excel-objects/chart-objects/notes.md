# MS Excel Objects

## Chart Objects

The [`ChartObject`](https://msdn.microsoft.com/en-us/VBA/Excel-VBA/articles/chartobject-object-excel) acts as a container for a [`Chart`](https://msdn.microsoft.com/en-us/VBA/Excel-VBA/articles/chart-object-excel#properties) on a given worksheet.

Reference any `ChartObject` or corresponding `Chart` through the `ChartObjects` collection:

```vb
Worksheets("Sheet1").ChartObjects.Count ' --> 2
Worksheets("Sheet1").ChartObjects(1).Name ' --> Chart 1
Worksheets("Sheet1").ChartObjects(1).Chart.Name ' --> Sheet 1 Chart 1
Worksheets("Sheet1").ChartObjects(1).Chart.HasTitle ' --> True
Worksheets("Sheet1").ChartObjects(1).Chart.ChartTitle.Text ' --> My Pie Chart
Worksheets("Sheet1").ChartObjects(1).Chart.ChartType ' --> 5 (number corresponds to a Pie Chart)
```

When using the `ChartType` property, reference this list of corresponding [Chart Types](https://msdn.microsoft.com/en-us/VBA/Excel-VBA/articles/xlcharttype-enumeration-excel).

Once you have learned about loops, you can loop through each chart in the collection:

```vb
Dim MyChart as ChartObject

For Each MyChart In Worksheets("Sheet1").ChartObjects
  MsgBox(MyChart.Name & " (" & MyChart.Chart.ChartType & ")")
Next MyChart
```
