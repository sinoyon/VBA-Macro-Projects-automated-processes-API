# VBA-Macro-Projects-automated-processes-API
This repository is a collection of my VBA projects and their related posts. VBA-Macro-Projects automated processes, UserForm, Windows Application Programming Interface(API).



VBA (Visual Basic for Applications) is a programming language that is built into Microsoft Office applications such as Excel, Word, and PowerPoint. VBA allows you to automate tasks within these applications, such as data manipulation, chart creation, and report generation.

Here is an example of VBA code that will automate the task of creating a new worksheet in an Excel workbook:

```
Sub CreateNewWorksheet()
    'Create a new worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "New Worksheet"
End Sub
```

This code uses the Add method of the Sheets collection to create a new worksheet in the active workbook. The After:= argument specifies that the new worksheet should be added after the last existing worksheet in the workbook. The code then assigns the name "New Worksheet" to the new worksheet.

Another example is that the following code will automate the task of moving data from one worksheet to another worksheet in the same workbook:


```
Sub MoveData()
    'Move data from Sheet1 to Sheet2
    Sheets("Sheet1").Range("A1:D10").Copy Destination:=Sheets("Sheet2").Range("A1")
End Sub
```

This code uses the Range object to select a range of cells on Sheet1 (A1:D10) and the Copy method to copy the selected range. The Destination:= argument specifies the location to where the data will be pasted, in this case Sheet2's range A1.

VBA code can be run from within the Office application by opening the VBA editor (pressing ALT + F11) and running the code by either running the subroutine directly or by placing the code in a button and then clicking the button.

VBA is a powerful tool for automating tasks in Office applications, and it can save a lot of time and effort compared to performing the same tasks manually. With VBA, you can create your own custom functions and macros to perform specific tasks, as well as automate repetitive tasks, or even create interactive forms and dialog boxes.

# Automating data sorting:

```
Sub SortData()
    'Sort data in range A1:D10 by column D in descending order
    Sheets("Sheet1").Range("A1:D10").Sort Key1:=Range("D1"), Order1:=xlDescending
End Sub
```

This code uses the Sort method to sort the data in the range A1:D10 by the values in column D in descending order. The Key1:= argument specifies the column to sort by, and the Order1:= argument specifies the sort order.

# Automating chart creation:

```
Sub CreateChart()
    'Create a chart showing sales data
    Dim chart As Chart
    Set chart = Sheets("Sheet1").Shapes.AddChart2(251, xlColumnClustered, Range("A1:D10")).Chart
    chart.ChartTitle.Text = "Sales Data"
End Sub
```


This code creates a new chart on Sheet1, using the data in the range A1:D10. The AddChart2 method is used to create the chart, with the 251 argument specifying the chart type (column clustered), and the Range("A1:D10") argument specifying the data range. The ChartTitle.Text property is used to set the chart title to "Sales Data".

# Automating data validation:

```
Sub DataValidation()
    'Add data validation to range A1:A10
    With Sheets("Sheet1").Range("A1:A10").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="Red, Green, Blue"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Invalid Input"
        .InputMessage = "Select a color from the list"
        .ErrorMessage = "Please select a color from the list"
        .ShowInput = True
        .ShowError = True
    End With
End Sub
```
This code adds data validation to the range A1:A10 on Sheet1, using the Validation property of the range. The Add method is used to add a validation rule of type xlValidateList, with the Formula1:="Red, Green, Blue" argument specifying the list of valid inputs. The other properties are used to customize the error messages and input prompts.

These are just a few examples of how you can use VBA to automate tasks in Excel. With a little creativity and effort, you can automate almost anything in Excel and improve your productivity. Remember to properly test your code and ensure that it works as expected before using it in a production environment, and also document your code to make sure it is readable and understandable for others.



</br></br>
👉 If you find this project useful, please ⭐ this repository 😆!</br></br>
👉 Check out my work on GitHub using similar data sets with SAS studio <a href="https://github.com/sinoyon?tab=repositories">Here. </a>

