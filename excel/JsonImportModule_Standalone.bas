'"
' & "JsonImportModule_Standalone.bas" & "
"

' Standalone JSON Parser Module for Excel
' This module converts JSON diagnostic tables to Excel worksheets with reserved offsets (5 rows, 2 columns), metadata display, and proper formatting.

Sub ImportJson()
    Dim jsonText As String
    Dim jsonData As Object
    Dim ws As Worksheet
    Dim jsonFile As String
    Dim jsonRow As Long

    ' Create a new worksheet for the JSON data
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Json_Import"

    ' Set reserved offsets
    For i = 1 To 5
        ws.Rows(i).Hidden = True
    Next i

    ' Read JSON file from user
    jsonFile = Application.GetOpenFilename("JSON Files (*.json), *.json")
    If jsonFile = "False" Then Exit Sub

    ' Read the JSON file\r\n    Open jsonFile For Input As #1
    jsonText = Input$(LOF(1), #1)
    Close #1

    ' Parse JSON text into a dictionary
    Set jsonData = JsonConverter.ParseJson(jsonText)

    ' Insert metadata
    ws.Cells(1, 1).Value = "Imported JSON Data"
    ws.Cells(2, 1).Value = "Source: " & jsonFile

    ' Populate worksheet with JSON data
    jsonRow = 7 ' Starting row for JSON data (after reserved offset)
    For Each item In jsonData
        ws.Cells(jsonRow, 1).Value = item("Key") ' Adjust key names as needed
        ws.Cells(jsonRow, 2).Value = item("Value") ' Adjust key names as needed
        jsonRow = jsonRow + 1
    Next item

    ' Format the worksheet
    ws.Columns.AutoFit
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Font.Size = 14
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 2)).Interior.Color = RGB(200, 200, 255)
End Sub

'