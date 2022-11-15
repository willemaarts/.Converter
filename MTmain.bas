Attribute VB_Name = "MTmain"
Sub MAIN_MT()

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
End With

Dim StartTime   As Double
Dim SecondsElapsed As Double
Dim Msg1 As Variant

If Range("C2").Value = "" Then
    MsgBox "Forecasting Excel MT not opened."
    With Application
        .ScreenUpdating = True
        .ScreenUpdating = True
    End With
    Exit Sub
End If

If Range("D2").Value = "" Then
    MsgBox "Forecasting Excel JUYO not opened."
    With Application
        .ScreenUpdating = True
        .ScreenUpdating = True
    End With
    Exit Sub
End If

StartTime = Timer

Dim wb As Workbook, wb1 As Workbook, wb2 As Workbook
Dim ws As Worksheet

On Error Resume Next

    On Error Resume Next
    Set wb = Workbooks("JUYO Forecasting formatter") '(.xlsm)
    
    If err.Number <> 0 Then
        If err.Number = 9 Then
            Debug.Print err.Number & " | " & err.Description
            Set wb = Workbooks("JUYO Forecasting formatter.xlsm") '(.xlsm)
            err.Clear
        End If
    End If
    
    Set ws = wb.Worksheets("Rekenblad")

wb.Activate
ws.Select

Debug.Print err.Number & " | " & err.Description

If err.Number <> 0 Then
    If err.Number = 9 Then
        MsgBox "Please make sure that the workbook is activated" & _
            "Please try again or contact the admin.", vbCritical
        Exit Sub
    End If
End If

Dim iMonth() As Variant, iMonthS As Variant
Dim iSegment() As Variant
Dim iData() As Variant, x As Long
Dim Y As Variant
Dim iNum As Integer, iNum1 As Integer, iNum2 As Integer
Dim match_quantity  As Variant, match_quantity1  As Variant, m As Variant

Dim emptyRow As Long, emptyRow1 As Long
Dim lastrow As Long
lastrow = Cells(Rows.count, "A").End(xlUp).Row

Y = Range("F2").Value

iMonth = Range("A2:A" & lastrow)
iMonthS = Range("A2").Value

For iNum = 1 To UBound(iMonth)
    Debug.Print iMonth(iNum, 1)
Next iNum

lastrow = Cells(Rows.count, "B").End(xlUp).Row

iSegment = Range("B2:B" & lastrow)

For iNum = 1 To UBound(iSegment)
    Debug.Print iSegment(iNum, 1)
Next iNum

c = UBound(iSegment) * 2

On Error Resume Next
Set wb1 = Workbooks(Range("C2").Value) 'error handling add

If err.Number <> 0 Then
    If err.Number = 9 Then
        Debug.Print err.Number & " | " & err.Description
        Set wb1 = Workbooks(Left(Range("C2").Value, Len(Range("C2").Value) - 5))
        err.Clear
    End If
End If

On Error Resume Next
Set wb2 = Workbooks(Range("D2").Value)

If err.Number <> 0 Then
    If err.Number = 9 Then
        Debug.Print err.Number & " | " & err.Description
        Set wb2 = Workbooks(Left(Range("D2").Value, Len(Range("D2").Value) - 5))
        err.Clear
    End If
End If

wb1.Activate

wb1.Unprotect

For iNum = 1 To UBound(iMonth)
    On Error Resume Next
    Set ws = wb1.Worksheets(iMonth(iNum, 1))
    
    Debug.Print err.Number & " | " & err.Description

    If err.Number <> 0 Then
        If err.Number = 9 Then
            MsgBox "Month: '" & iMonth(iNum, 1) & "'. Is not recognised in as sheet in: " & wb1.Name & vbNewLine & _
                vbNewLine & "Please try again or contact the admin.", vbCritical
            Exit Sub
        End If
    End If

    ws.Visible = xlSheetVisible
    ws.Select
    
    If Y > 1 Then
        If Year(Range(Cells(4, 4), Cells(4, 4)).Value) = Y Then
            Debug.Print Y & " = first year of: " & iMonth(iNum, 1) & Year(Range(Cells(4, 4), Cells(4, 4)).Value)
            Y = 0
        Else
            MsgBox "Selected year: " & Y & ", does not correspond with first year of: " & iMonth(iNum, 1) & vbNewLine & vbNewLine & _
                "Please contact the admin.", vbCritical
            Exit Sub
        End If
    End If
        
    For iNum1 = 1 To UBound(iSegment)
        match_quantity1 = WorksheetFunction.Match("Group Total", Range("C:C"), 0)
        
        On Error Resume Next
        match_quantity = WorksheetFunction.Match(iSegment(iNum1, 1), Range("C200:C" & match_quantity1), 0)
        
        Debug.Print err.Number & " | " & err.Description
        
        If err.Number <> 0 Then
            If err.Number = 1004 Then
                MsgBox "Segment: '" & iSegment(iNum1, 1) & "'. Is not recognised in: '" & iMonth(iNum, 1) & "'." & vbNewLine & _
                    vbNewLine & "Please try again and upload the segments from the forecast file.", vbCritical
                Exit Sub
            End If
        End If
        
        Debug.Print "Month: " & iMonth(iNum, 1) & " | Segment: " & iSegment(iNum1, 1) & " | in row: " & _
            match_quantity + 199 & " of " & match_quantity1

        Select Case iMonth(iNum, 1)
        
            Case "Jan Fcst", "Mar Fcst", "May Fcst", "Jul Fcst", "Aug Fcst", "Oct Fcst", "Dec Fcst": m = 31 + 3
            Case "Apr Fcst", "Jun Fcst", "Sep Fcst", "Nov Fcst": m = 30 + 3
            Case "Feb Fcst": m = 28 + 3
            Case Else
                Debug.Print "No Month"
                m = 31 + 3
        
        End Select

        ReDim Preserve iData(x)
        iData(x) = Application.WorksheetFunction.Transpose(Range(Cells(match_quantity + 202, 4), Cells(match_quantity + 202, m)))
        
        x = x + 1
        
        ReDim Preserve iData(x)
        iData(x) = Application.WorksheetFunction.Transpose(Range(Cells(match_quantity + 208, 4), Cells(match_quantity + 208, m)))
        
        x = x + 1
            
    Next iNum1
 
Next iNum

wb.Activate
Sheets("OutP").Select

x = 0
l = 1
    For x = 0 To UBound(iData)
        Range(Cells(1, l), Cells(UBound(iData), l)) = iData(x)
        l = l + 1
    Next x

With Sheets("OutP")
    .Rows(32 & ":" & .Rows.count).Delete
End With

Cells.Replace what:="#N/A", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
        , FormulaVersion:=xlReplaceFormula2

Range("A1").Select

For x = 1 To UBound(iMonth) - 1
    emptyRow1 = WorksheetFunction.CountA(Range("AA:AA")) + 1
    emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1

    Range(Cells(emptyRow, 1), Cells(emptyRow + emptyRow1, c)).Value = Range(Cells(1, c + 1), Cells(emptyRow1 - 1, c * 2)).Value
    Range(Cells(1, c + 1), Cells(1, c * 2)).EntireColumn.Delete
    
    Cells.Replace what:="#N/A", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
        , FormulaVersion:=xlReplaceFormula2

Next x

Dim sourceColumn As Range, targetColumn As Range

emptyRow = WorksheetFunction.CountA(Range("A:A"))
Range("A1:AN" & emptyRow).Select

Set sourceColumn = wb.ActiveSheet.Range("A1:AN" & emptyRow)
Set targetColumn = wb2.Worksheets(1).Range("B2")

Selection.Copy Destination:=targetColumn

Cells.Select
Selection.ClearContents

Select Case iMonthS
    Case "Jan Fcst": iMonthS = "1/1/2022"
    Case "Feb Fcst": iMonthS = "2/1/2022"
    Case "Mar Fcst": iMonthS = "3/1/2022"
    Case "Apr Fcst": iMonthS = "4/1/2022"
    Case "May Fcst": iMonthS = "5/1/2022"
    Case "Jun Fcst": iMonthS = "6/1/2022"
    Case "Jul Fcst": iMonthS = "7/1/2022"
    Case "Aug Fcst": iMonthS = "8/1/2022"
    Case "Sep Fcst": iMonthS = "9/1/2022"
    Case "Oct Fcst": iMonthS = "10/1/2022"
    Case "Nov Fcst": iMonthS = "11/1/2022"
    Case "Dec Fcst": iMonthS = "12/1/2022"
End Select

wb2.Activate
With Range("A2")
    .Select
    .NumberFormat = "yyyy-mm-dd;@"
    .FormulaR1C1 = iMonthS
    .AutoFill Destination:=Range("A2:A" & emptyRow + 1)
End With

wb.Activate

SecondsElapsed = Round(Timer - StartTime, 6)
Debug.Print "Ran in: " & SecondsElapsed & " seconds"

Sheets("Rekenblad").Select

With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
End With

Msg1 = "Forcasting file succesfully formatted to: " & wb2.Name & " file, in: " & SecondsElapsed & " seconds."

MsgBox Msg1, vbExclamation, "Formatted succesfully"

Application.ScreenUpdating = True


End Sub

Sub StartConvert()

MTprocess.Show

End Sub
