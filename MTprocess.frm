VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MTprocess 
   Caption         =   "UserForm2"
   ClientHeight    =   8805.001
   ClientLeft      =   420
   ClientTop       =   1065
   ClientWidth     =   8775.001
   OleObjectBlob   =   "MTprocess.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MTprocess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdConvert_Click()
    
    Dim msg              As Variant
    
    msg = ("Segments are not evenly distributed! Make sure that the segments in both listboxes are 100% correct. " & vbNewLine & vbNewLine & _
          "Segments JUYO count  : " & Me.ListBox2.ListCount & vbNewLine & _
          "Segments Client count : " & Me.ListBox3.ListCount)
    
    If Me.ListBox2.ListCount = Me.ListBox3.ListCount Then
        Debug.Print "COUNT Segments: " & Me.ListBox2.ListCount & " | " & Me.ListBox3.ListCount
    Else
        MsgBox msg, vbCritical, "Segments are Not correct!"
        Exit Sub
    End If
    
    If CheckBox2.Value = True Then
        Call CmdStoreSegments_Click
    End If
   
    Unload Me
    Call MAIN_MT
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
End Sub

Private Sub CmdLastUsedSeg_Click()
    Dim iMonth()         As Variant
    Dim iNum             As Integer
    Dim lastrow          As Long
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
    ListBox3.Clear
    
    lastrow = Cells(Rows.count, "B").End(xlUp).Row
    
    iMonth = Range("B2:B" & lastrow)
    
    For iNum = 1 To UBound(iMonth)
        Me.ListBox3.AddItem iMonth(iNum, 1)
    Next iNum
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
End Sub

Private Sub CmdLoadSegments_Click()
    Dim iVarSeg()        As Variant, x As Long, Y As Long
    Dim iNum             As Integer
    Dim match_quantity   As Variant, match_quantity1  As Variant, m As Variant
    Dim emptyRow         As Long, CB As Long
    
    Dim wb               As Workbook, wb1 As Workbook
    Dim ws               As Worksheet, ws1 As Worksheet
    
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
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
    ListBox3.Clear
    
    wb.Activate
    ws.Select
    
    On Error Resume Next
    Set wb1 = Workbooks(Range("C2").Value)
    
    If err.Number <> 0 Then
        If err.Number = 9 Then
            Debug.Print err.Number & " | " & err.Description
            Set wb1 = Workbooks(Left(Range("C2").Value, Len(Range("C2").Value) - 5))
            err.Clear
        End If
    End If
    
    Set ws1 = wb1.Worksheets(Range("A2").Value) 'Maybe also use last segment to see if they differ?
    
    wb1.Activate
    wb1.Unprotect
    ws1.Visible = xlSheetVisible
    ws1.Select
    
    match_quantity1 = WorksheetFunction.Match("Total Rooms BOB", Range("C:C"), 0) - 1        '400
    match_quantity = WorksheetFunction.Match("ROOMS REVENUE BY SEGMENT", Range("C:C"), 0)        '202
    
    CB = WorksheetFunction.CountBlank(Range(Cells(match_quantity, 3), Cells(match_quantity1, 3))) + 1        'count whiterows
    Debug.Print "White rows in CB: " & CB
    m = match_quantity1 - match_quantity - 14 - CB
    
    m = m / 12        '# of segments
    
    For Y = 1 To m
        If Range(Cells(match_quantity + 2, 3), Cells(match_quantity + 2, 3)).Value = "Transient Total" Then
            match_quantity = match_quantity + 8
        End If
        ReDim Preserve iVarSeg(x)
        iVarSeg(x) = Range(Cells(match_quantity + 2, 3), Cells(match_quantity + 2, 3)).Value
        match_quantity = match_quantity + 12
        x = x + 1
    Next Y
    
    For iNum = 0 To UBound(iVarSeg)
        Me.ListBox3.AddItem iVarSeg(iNum)
    Next iNum
    
    wb.Activate
    ws.Select
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
End Sub

Private Sub CmdLoad_Click()        'Excel files load
    
    Application.ScreenUpdating = False
    
    Dim wb               As Workbook, wb1 As Workbook, wb2 As Workbook
    Dim ws               As Worksheet, ws1 As Worksheet
    
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
    
    If ComboBox1.Value = "" Then
        MsgBox "No name selected!"
        Exit Sub
    End If
    
    If ComboBox2.Value = "" Then
        MsgBox "No name selected!"
        Exit Sub
    End If
    
    Range("C2").Value = ComboBox1.Value
    Range("D2").Value = ComboBox2.Value
    
    Dim iVarSeg()        As Variant, x As Long, Y As Long
    Dim iNum             As Integer
    Dim match_quantity   As Variant, match_quantity1  As Variant, m As Variant
    Dim emptyRow         As Long, CB As Long
    
    On Error Resume Next
    Set wb1 = Workbooks(Range("D2").Value)
    
    If err.Number <> 0 Then
        If err.Number = 9 Then
            Debug.Print err.Number & " | " & err.Description
            Set wb1 = Workbooks(Left(Range("D2").Value, Len(Range("D2").Value) - 5))
            err.Clear
        End If
    End If
    
    Set ws1 = wb1.Worksheets("Sheet0")
    
    wb1.Activate
    ws1.Select
    
    Dim iSegment()       As Variant
    Dim lastColumn       As Long
    
    lastColumn = Range("A1").End(xlToRight).Column
    
    iSegment = Application.WorksheetFunction.Transpose(Range(Cells(1, 1), Cells(1, lastColumn)))
    
    For iNum = 2 To UBound(iSegment) Step 2
        Debug.Print Left(iSegment(iNum, 1), Len(iSegment(iNum, 1)) - 3)
        Me.ListBox2.AddItem Left(iSegment(iNum, 1), Len(iSegment(iNum, 1)) - 3)
    Next iNum
    
    Me.Height = 196
    
    wb.Activate
    ws.Select
    
    If CheckBox1.Value = True Then
        Call CmdLastUsedSeg_Click
    Else
        Call CmdLoadSegments_Click
    End If
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub CmdStoreSegments_Click()        'Store segments
    
    Application.ScreenUpdating = False
    
    Dim x                As Integer
    Dim lastrow          As Long
    
    lastrow = Cells(Rows.count, "B").End(xlUp).Row
    
    Range("B2:B" & lastrow).ClearContents
    Range("B2").Select
    
    For x = 0 To Me.ListBox3.ListCount - 1
        Me.ListBox3.Selected(x) = True
        If Me.ListBox3.Selected(x) = True Then
            ActiveCell = Me.ListBox3.List(x)
            ActiveCell.Offset(1, 0).Select
        End If
    Next x
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub CmdMonths_Click()
    
    Application.ScreenUpdating = False
    
    Dim m As Integer, m1 As Integer, z As Variant 'months
    Dim Y As Integer 'years
    Dim x As Integer, x1 As Integer 'For loop
    Dim lastrow As Long
    
    If CBm1.Value = "" Then
        MsgBox "No First Month Selected."
        Exit Sub
    End If
    
    If CBm2.Value = "" Then
        MsgBox "No End Month Selected."
        Exit Sub
    End If
    
    If CBy1.Value = "" Then
        MsgBox "No Year Selected."
        Exit Sub
    End If
    
    lastrow = Cells(Rows.count, "A").End(xlUp).Row + 1
    
    Range("A2:A" & lastrow).ClearContents
    
    lastrow = Cells(Rows.count, "A").End(xlUp).Row + 1
    
    Y = CBy1.Value
    m = CBm1.Value
    m1 = CBm2.Value
    
    Range(Cells(2, 6), Cells(2, 6)).Value = Y
    
    Select Case m
        Case Is = m1: x1 = 1
        Case Is < m1: x1 = 1 + (m1 - m)
        Case Is > m1: x1 = (12 + 1 - m) + (m1)
        Case Else:
            x1 = 1
            Debug.Print "Error; calculation with months."
    End Select
    
    For x = 1 To x1
        
        Select Case m
            Case "1": z = "Jan Fcst"
            Case "2": z = "Feb Fcst"
            Case "3": z = "Mar Fcst"
            Case "4": z = "Apr Fcst"
            Case "5": z = "May Fcst"
            Case "6": z = "Jun Fcst"
            Case "7": z = "Jul Fcst"
            Case "8": z = "Aug Fcst"
            Case "9": z = "Sep Fcst"
            Case "10": z = "Oct Fcst"
            Case "11": z = "Nov Fcst"
            Case "12": z = "Dec Fcst"
        End Select
        
        Debug.Print m & " " & Y & " " & z
    
        Range(Cells(lastrow, 1), Cells(lastrow, 1)).Value = z
        lastrow = lastrow + 1
        
        m = m + 1
        
        If m = 12 + 1 Then
            m = m - 12
            Y = Y + 1
        Else
            m = m
        End If
    
    Next x
    
    With Me
        .Height = 445
        .Top = 170
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub UserForm_Initialize()
    
    Dim iMonth()         As Variant
    Dim iSegment()       As Variant
    Dim iNum             As Integer, iNum1 As Integer, iNum2 As Integer
    Dim vWorkbook        As Workbook
    
    Application.ScreenUpdating = False
    
    For x = 0 To 3
        Me.CBy1.AddItem Year(Now()) + x
    Next x
    
    For x = 1 To 12
        Me.CBm1.AddItem x
        Me.CBm2.AddItem x
    Next x
    
    ComboBox1.Clear
    ComboBox2.Clear
    
    For Each vWorkbook In Workbooks
        ComboBox1.AddItem vWorkbook.Name
        ComboBox2.AddItem vWorkbook.Name
    Next
    
    Dim wb               As Workbook
    Dim ws               As Worksheet
    
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets("Rekenblad")
    
    wb.Activate
    ws.Select
    
    Range("C2").Value = ""
    Range("D2").Value = ""
    
    With Me
        .Height = 113
        .Width = 390
    End With
    
    CheckBox1.Value = True
    
    Application.ScreenUpdating = True
    
End Sub

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'----------------------------FUNCTION KEYS--------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

Private Sub CmdRight_Click()
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
    Dim x                As Integer
    
    For x = 0 To Me.ListBox3.ListCount - 1
        If Me.ListBox3.Selected(x) = True Then
            ListBox4.AddItem Me.ListBox3.List(x)
            ListBox3.RemoveItem x
        End If
    Next x
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
End Sub
Private Sub CmdUp_Click()
    
    Application.ScreenUpdating = False
    
    With Me.ListBox3
        
        If .ListIndex = -1 Then Exit Sub
        If .ListIndex = 0 Then Exit Sub
        
        curIndex = .ListIndex
        
        curVal = .List(curIndex)
        othIndex = curIndex - 1
        othVal = .List(othIndex)
        
        .List(othIndex) = curVal
        .List(curIndex) = othVal
        
        .Selected(othIndex) = True
        
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub CmdDown_Click()
    
    Application.ScreenUpdating = False
    
    With Me.ListBox3
        
        If .ListIndex = -1 Then Exit Sub
        If .ListIndex = .ListCount - 1 Then Exit Sub
        
        curIndex = .ListIndex
        
        curVal = .List(curIndex)
        othIndex = curIndex + 1
        othVal = .List(othIndex)
        
        .List(othIndex) = curVal
        .List(curIndex) = othVal
        
        .Selected(othIndex) = True
        
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub CmdRight_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call CmdRight_Click
End Sub

Private Sub CmdLeft_Click()
    
    Application.ScreenUpdating = False
    
    'Move selected items to the Left
    
    'Remove item from the left and place it on the right.
    'Loop through the items
    For itemIndex = ListBox4.ListCount - 1 To 0 Step -1
        
        'Check if an item was selected.
        If ListBox4.Selected(itemIndex) Then
            
            'Move selected item to the right.
            ListBox3.AddItem ListBox4.List(itemIndex)
            
            'Remove selected item from the left.
            ListBox4.RemoveItem itemIndex
            
        End If
        
    Next itemIndex
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub CmdLeft_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call CmdLeft_Click
End Sub

