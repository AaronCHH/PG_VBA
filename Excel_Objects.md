# Excel Objects
- [Excel Objects](#Excel-Objects)
  - [Introduction](#Introduction)
  - [Application Objects](#Application-Objects)
  - [Workbook Objects](#Workbook-Objects)
  - [Worksheet Objects](#Worksheet-Objects)
  - [Range Objects](#Range-Objects)
  - [Reference](#Reference)

## Introduction

When programming using VBA, there are few important objects that a user would be dealing with.

* Application Objects
* Workbook Objects
* Worksheet Objects
* Range Objects

## Application Objects
The Application object consists of the following −

Application-wide settings and options.
Methods that return top-level objects, such as ActiveCell, ActiveSheet, and so on.

__Example__
```
'Example 1 :
Set xlapp = CreateObject("Excel.Sheet") 
xlapp.Application.Workbooks.Open "C:\test.xls"
```

```
'Example 2 :
Application.Windows("test.xls").Activate
```

```
'Example 3:
Application.ActiveCell.Font.Bold = True
```

## Workbook Objects
The Workbook object is a member of the Workbooks collection and contains all the Workbook objects currently open in Microsoft Excel.

__Example__
```
'Ex 1 : To close Workbooks
Workbooks.Close
```
```
'Ex 2 : To Add an Empty Work Book
Workbooks.Add
```
```
'Ex 3: To Open a Workbook
Workbooks.Open FileName:="Test.xls", ReadOnly:=True
```
```
'Ex : 4 - To Activate WorkBooks
Workbooks("Test.xls").Worksheets("Sheet1").Activate
```

## Worksheet Objects
The Worksheet object is a member of the Worksheets collection and contains all the Worksheet objects in a workbook.

__Example__
```
'Ex 1 : To make it Invisible
Worksheets(1).Visible = False
```
```
'Ex 2 : To protect an WorkSheet
Worksheets("Sheet1").Protect password:=strPassword, scenarios:=True
```

## Range Objects
Range Objects represent a cell, a row, a column, or a selection of cells containing one or more continuous blocks of cells.
```
'Ex 1 : To Put a value in the cell A5
Worksheets("Sheet1").Range("A5").Value = "5235"
```
```
'Ex 2 : To put a value in range of Cells
Worksheets("Sheet1").Range("A1:A4").Value = 5
```

## Reference
* Tutorialspoints https://www.tutorialspoint.com/vba/vba_excel_objects.htm