Attribute VB_Name = "search"
Sub FindName()
    Dim targetWorkbook As Workbook
    Dim ws As Worksheet
    Dim cell As Range

    Dim sourceFolder As String
    Dim Filename As String
    Dim resultRow As Long, resultCol As Long
    Dim targetCellContent As String
    Dim targetWorksheet As Worksheet
    Dim searchRange As Range
    Dim foundRange As Range
    Dim firstFoundCell As Range
    Dim startPos As Long, Length As Long

    sourceFolder = ThisWorkbook.Sheets(1).Cells(2, 1).Value
    resultRow = 3
    resultCol = 5

    Filename = Dir(sourceFolder & "\*.xlsx")
    
    Do While Filename <> ""
        Dim sourceWorksheet As Worksheet
        Set sourceWorksheet = ThisWorkbook.Sheets("Sheet1")
        Set targetWorkbook = Workbooks.Open(sourceFolder & "\" & Filename)
        
        For Each ws In targetWorkbook.Worksheets
            If ws.Name <> "ï\éÜ" And ws.Name <> "ïœçXóöó" Then
                For i = 2 To sourceWorksheet.Cells(Rows.Count, 2).End(xlUp).Row
                    targetCellContent = sourceWorksheet.Cells(i, 2).Value
                    Set targetWorksheet = ws
                    Set searchRange = targetWorksheet.UsedRange
                    Set foundRange = searchRange.Find(What:=targetCellContent, LookIn:=xlValues, LookAt:=xlPart)
                    Set firstFoundCell = foundRange
                    
                    If Not foundRange Is Nothing Then
                        Do
                            sourceWorksheet.Cells(resultRow, resultCol).Value = foundRange.Value
                            sourceWorksheet.Cells(resultRow, resultCol + 1).Value = targetWorkbook.Name
                            sourceWorksheet.Cells(resultRow, resultCol + 2).Value = ws.Name & ",row:" & foundRange.Row & ",col:" & foundRange.Column
                            sourceWorksheet.Cells(resultRow, resultCol + 3).Value = sourceFolder & "\" & Filename

                            startPos = InStr(foundRange.Value, targetCellContent)
                            Length = Len(targetCellContent)
                            With sourceWorksheet.Cells(resultRow, resultCol).Characters(startPos, Length).Font
                                .Color = RGB(255, 0, 0)
                                .Bold = True
                            End With
                            
                            resultRow = resultRow + 1
                            Set foundRange = searchRange.FindNext(foundRange)
                        Loop Until foundRange Is Nothing Or foundRange.Address = firstFoundCell.Address
                    End If
                Next i
            End If
        Next ws
        
        targetWorkbook.Close SaveChanges:=True
        Filename = Dir()
    Loop

    Set targetWorkbook = Nothing
End Sub

