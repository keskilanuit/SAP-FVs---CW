Sub Load_FV60s()

    Dim sourceFolder As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim entryWs As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim cell As Range
    
        With ThisWorkbook.Sheets("Entry").Range("A10:AE999")
        .Value = vbNullString
        .ClearFormats
    End With
    'loader files from path located on cell C1
sourceFolder = ThisWorkbook.Sheets("Entry").Range("C1").Value
Set entryWs = ThisWorkbook.Sheets("Entry")
        'load all files from source folder if .xlsx file contains text "FV60" in file names'
Dim filename As String
filename = Dir(sourceFolder & "*FV60*.xlsx")
Do While filename <> ""

    Application.DisplayAlerts = True ' Turn off display alerts

    Set wb = Workbooks.Open(sourceFolder & filename, UpdateLinks:=False)

    For Each ws In wb.Worksheets
        lastRow = entryWs.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To ws.UsedRange.Rows.Count
            For j = 1 To ws.UsedRange.Columns.Count
                entryWs.Cells(lastRow + i - 1, j).Value = ws.Cells(i, j).Value
            Next j
        Next i
    Next ws

    wb.Close False ' Close without saving changes
    Application.DisplayAlerts = True ' Turn on display alerts

     counter = counter + 1 ' Increment counter
     
    filename = Dir
Loop
    Range("E10:E999").NumberFormat = "MM/DD/YYYY"
    Range("G10:G999").NumberFormat = "MM/DD/YYYY"
    Range("M10:M999").NumberFormat = "MM/DD/YYYY"
    
'loop through value in column Y, if any cells contains value, round value into 2 decimal places'
    For Each cell In Range("Y10:Y999")
        If Not IsEmpty(cell.Value) Then
            cell.Value = Round(cell.Value, 2)
        End If
    Next cell
    
 
    
    'if range within column Y contains value of 0, then entire row delete.'
    
           Count = Worksheets("Entry").Cells(Rows.Count, "Y").End(xlUp).Row
        i = 10
    Do While i <= Count
    If Cells(i, 25) = "0" Then
    Rows(i).EntireRow.Delete
    i = i - 1
    End If
    i = i + 1
    Count = Worksheets("Entry").Cells(Rows.Count, "Y").End(xlUp).Row
    
    Loop
    
        lastRow = Cells(Rows.Count, "B").End(xlUp).Row 'find the last row of column B
    
    For i = 10 To Cells(Rows.Count, "B").End(xlUp).Row
        If Cells(i, "B").Value = "1" Then
            Cells(i, "B").Value = "'001"
            Cells(i, "B").NumberFormat = "@"
        End If

            'For DC FV Company Codes only'
        If Cells(i, "B").Value = "66" Then
         Cells(i, "B").Value = "'066"
        Cells(i, "B").NumberFormat = "@"
    Next i

    MsgBox counter & " file(s) from folder has been successfully loaded into the Master Template!", vbInformation, "FV60 Loader"
    

    
End Sub
