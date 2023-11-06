Sub DummyCMSREAD()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long, j As Integer
    Dim cellValue As String, remainder As String
    Dim answer As VbMsgBoxResult

    ' Speed up by turning off screen updating and calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("Dummy")

    ' Find the last row with data in Column A
    LastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    For i = 2 To LastRow
        Application.StatusBar = "Processing row " & i & " of " & LastRow

        
        If IsEmpty(ws.Cells(i, 1)) Then Exit For
 ' Process Column B
        cellValue = ws.Cells(i, 2).Value
        If AlphanumericCount(cellValue) <> 2 Then
            ws.Cells(i, 2).Value = Left(OnlyAlphanumeric(cellValue), 2)
        End If

        ' Process Column C
        cellValue = ws.Cells(i, 3).Value
        If AlphanumericCount(cellValue) > 6 Then
            remainder = Mid(OnlyAlphanumeric(cellValue), 7)
            ws.Cells(i, 3).Value = Left(OnlyAlphanumeric(cellValue), 6)
            ws.Cells(i, 4).Value = Trim(remainder & " " & ws.Cells(i, 4).Value)
        End If

        ' Process Column D
        cellValue = ws.Cells(i, 4).Value
        If AlphanumericCount(cellValue) > 3 Then
            remainder = Mid(OnlyAlphanumeric(cellValue), 4)
            ws.Cells(i, 4).Value = Left(OnlyAlphanumeric(cellValue), 3)
            ws.Cells(i, 5).Value = Trim(remainder & " " & ws.Cells(i, 5).Value)
        End If

        ' Process Column E
        cellValue = OnlyAlphanumeric(ws.Cells(i, 5).Value)
        If Len(cellValue) > 6 Then
            If Not IsEmpty(ws.Cells(i, 4).Value) Then
                answer = MsgBox("Row " & i & " has data in Column D. Do you wish to proceed?", vbYesNo)
                If answer = vbNo Then GoTo NextRow
            End If

            ws.Cells(i, 4).Value = Right(cellValue, Len(cellValue) - 6)
            ws.Cells(i, 5).Value = Left(cellValue, 6)
        End If

NextRow:
    Next i
    
    ' Reset everything
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
End Sub

' Function to count only alphanumeric characters
Function AlphanumericCount(str As String) As Integer
    Dim i As Integer, count As Integer
    count = 0
    For i = 1 To Len(str)
        If Mid(str, i, 1) Like "[0-9A-Za-z]" Then
            count = count + 1
        End If
    Next i
    AlphanumericCount = count
End Function

' Function to return only alphanumeric characters from a string
Function OnlyAlphanumeric(str As String) As String
    Dim i As Integer, result As String
    result = ""
    For i = 1 To Len(str)
        If Mid(str, i, 1) Like "[0-9A-Za-z]" Then
            result = result & Mid(str, i, 1)
        End If
    Next i
    OnlyAlphanumeric = result
End Function

Sub CopyTwoDigitCode()

    Dim ws As Worksheet
    Dim lastRowA As Long
    Dim i As Long
    Dim twoDigitCode As String

    ' Speed up by turning off screen updating and calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Assume the active sheet is where the data is
    Set ws = ThisWorkbook.ActiveSheet

    ' Find the last row with data in Column A
    lastRowA = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' Start from B2
    i = 2

    Do While i <= lastRowA
        Application.StatusBar = "Processing row " & i & " of " & lastRowA

        ' Check if the value in Column B is a 2-digit code
        If Len(ws.Cells(i, 2).Value) = 2 Then
            twoDigitCode = ws.Cells(i, 2).Value
            
            ' Check the cell below in Column C for a 6-digit code
            i = i + 1
            Do While Len(ws.Cells(i, 3).Value) = 6
                ws.Cells(i, 2).Value = twoDigitCode
                i = i + 1
            Loop
        Else
            i = i + 1
        End If
        
    Loop

    ' Reset everything
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False

End Sub

Sub CheckAndDeleteEmptyColumns()

    Dim ws As Worksheet
    Dim lastCol As Long, lastRowA As Long, i As Long, r As Long
    Dim emptyCols As String, nonEmptyCols As String

    ' Speed up by turning off screen updating and calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Assume the active sheet is where the data is
    Set ws = ThisWorkbook.ActiveSheet

    ' Find the last column with data in Row 1 starting from Column N
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    lastRowA = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' Start the search and deletion from column N onwards
    For i = 14 To lastCol ' 14 corresponds to column N
        Application.StatusBar = "Processing column " & i & " of " & lastCol

        Dim foundData As Boolean
        foundData = False
        
        For r = 2 To lastRowA
            If ws.Cells(r, i).Value <> "" Then
                ws.Cells(r, i).Select ' Select the cell with data
                foundData = True
                nonEmptyCols = nonEmptyCols & " " & Split(ws.Cells(1, i).Address, "$")(1) & ","
                Exit For
            End If
        Next r

        If foundData Then
            MsgBox "Data was found in the column:" & nonEmptyCols & " The first cell with data has been selected."
            Exit Sub
        Else
            emptyCols = emptyCols & " " & Split(ws.Cells(1, i).Address, "$")(1) & ","
        End If
    Next i

    ' Delete empty columns
    If emptyCols <> "" Then
        For i = lastCol To 14 Step -1
            If InStr(1, emptyCols, Split(ws.Cells(1, i).Address, "$")(1)) > 0 Then
                ws.Columns(i).Delete
            End If
        Next i
        MsgBox "All identified columns (" & emptyCols & ") were empty and have been deleted."
    End If

    ' Process columns H-L
    Dim highlightRange As Range
    Set highlightRange = ws.Range("H2:L" & lastRowA)
    For Each cell In highlightRange
        If IsNumeric(cell.Value) Then
            cell.Value = cell.Value * 1
        End If
    Next cell

    ' Reset everything
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False

End Sub

Sub FilterFormulaNAandExcludeBlanks()

    Dim ws As Worksheet
    Dim lastRowA As Long
    Dim rng As Range

    ' Speed up by turning off screen updating and calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Assume the active sheet is where the data is
    Set ws = ThisWorkbook.ActiveSheet

    ' Find the last row with data in Column A
    lastRowA = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' Define the range to apply the filter to (assuming up to column Z for the sake of example)
    Set rng = ws.Range("A1:E" & lastRowA)

    ' Clear any existing filters
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    ' Apply filter to Column A to show only cells with errors
    rng.AutoFilter Field:=1, Criteria1:="=#N/A"

    ' Now, filter column E to exclude blanks
    rng.AutoFilter Field:=5, Criteria1:="<>" & ""

    ' Reset screen updating and calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

