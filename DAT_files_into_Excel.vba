' ==============================================================================
' Project:      Bulk .DAT File Importer
' Author:       Krisztián Princzinger
' Description:  Reads and consolidates multitudes of .DAT files into a single sheet.
'               Parses semicolon-delimited data and enforces data types.
' Note:         This was one of my first automation projects, designed to solve
'               an immediate business need for mass data consolidation.
'               It was originally written with Hungarian variables and comments,
'               but I refactored them for this upload.
' License:      MIT License
' ==============================================================================

Option Explicit

Sub ImportAndConsolidateFiles()

    Dim folderPath As String
    Dim fileName As String
    Dim textLine As String
    Dim nextRow As Long
    Dim ws As Worksheet
    Dim dataParts As Variant
    Dim i As Long
    Dim fileNum As Integer

    'Declare the target worksheet, with creating one if it doesn't exist
    On Error Resume Next
        Set ws = ThisWorkbook.Sheets("Raw Data(VBA)")
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Sheets.Add
            ws.Name = "Raw Data(VBA)"
        Else
            ws.Cells.Clear
        End If
    On Error GoTo 0

    'Folder selection
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the folder containing the .DAT files"
        If .Show <> -1 Then Exit Sub 'Single line IF syntax in VBA
        folderPath = .SelectedItems(1) & "\"
    End With

    'File loop
    fileName = Dir(folderPath & "*.DAT")

    Do While fileName <> ""
        fileNum = FreeFile()
        Open folderPath & fileName For Input As #fileNum

        'Read file line by line
        Do Until EOF(fileNum)
            Line Input #fileNum, textLine
            dataParts = Split(textLine, ";")
        
            nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

            'Data Type Enforcement & Import
            For i = 0 To UBound(dataParts)
                Select Case i
                    Case 1, 8 'Column 2 and 9 as Text
                        ws.Cells(nextRow, i + 1).Value = "'" & dataParts(i)
                    Case 5, 7 'Column 6 and 8 as Numbers
                        If IsNumeric(dataParts(i)) Then
                            ws.Cells(nextRow, i + 1).Value = CLng(dataParts(i))
                        Else
                            ws.Cells(nextRow, i + 1).Value = dataParts(i)
                        End If
                    Case Else
                        ws.Cells(nextRow, i + 1).Value = dataParts(i)
                End Select
            Next i
        Loop
        Close #fileNum
        fileName = Dir
    Loop
    MsgBox "Process Complete!", vbInformation
End Sub

