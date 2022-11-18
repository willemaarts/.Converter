VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MTprocess 
   Caption         =   "UserForm2"
   ClientHeight    =   672.75
   ClientLeft      =   15
   ClientTop       =   30
   ClientWidth     =   388.5
   OleObjectBlob   =   "init.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "init"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' wb = Converter file | wb1 = Client | wb2 = Juyo
Public wb As Workbook, wb1 As Workbook, wb2 as Workbook
Public ws As Worksheet, ws1 As Worksheet, ws2 as Workbook
Public rng As Range
Public cel As Range
Public Y as Integer

Public fullRange As Range, fullRange1 as range

Const err1 As Variant = vbNewLine & vbNewLine & _
                        "Workbook name is not the same. Please try again."

' Ctrl + '/' to quicly comment

Private Sub UserForm_Initialize()

    Dim vWorkbook As Workbook
    dim x as Integer
    
    If Range("E1").value <> "EXCEL FILE" Then
        MsgBox "Please make sure that the converter file is active (by clicking on it)"
        me.hide
        exit sub
    end if

    ' Gets the current Workbook name
    Range("E2").Value = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)

    ' This will clear the combobox values of the last Workbook names used.
    fileC.Clear
    fileJ.Clear

    ' Will get all the open workbooks names and put them in the ComboBoxes.
    For Each vWorkbook In Workbooks
        If vWorkbook.Name <> ThisWorkbook.Name Then
            fileC.AddItem vWorkbook.Name
            fileJ.AddItem vWorkbook.Name
        End If
    Next
    
    ' This will get the Converter tool name and store it as a variable
    On Error Resume Next
    Set wb = Workbooks(Range("E2").Value)
    
    If err.Number <> 0 Then
        If err.Number = 9 Then
            MsgBox err.Number & " | " & err.Description & err1
            debug.print "3"
            err.Clear
            Exit Sub
        End If
    End If
    
    Set ws = wb.Worksheets("Rekenblad")
    
    wb.Activate
    ws.Select

    for x = 0 to 3
        me.yearBx.AddItem Year(Now()) + x
    next x

    yearBx.ListIndex = 0
    
    ' Will clear the last used workbook names.
    Range("C2").Value = ""
    Range("D2").Value = ""
    
    ' //TODO Can deleted after debuggin is done
    fileC.ListIndex = 0
    fileJ.ListIndex = 1

    me.label13.Top =  1000
    
    ' This will disable, so that the users can't click anything else.
    ' Now everything is in the right order
    CmdLoad.Enabled = True
    CmdSheets.Enabled = False
    segmentbx.Enabled = False
    CmdSegments.Enabled = False
    cmdTerminology.Enabled = False
    CmdStoreSegments.Enabled = False
    CmdLastUsedSeg.Enabled = False
    OptionButton1.Enabled = False
    OptionButton2.Enabled = False
    
End Sub

Private Sub CmdLoad_Click() 'This will start the first process.

    Dim lastColumn As Long
    Dim iNum As Integer
    Dim iSegment() As Variant

    wb.Activate
    ws.Select
    
    If fileC.Value = "" Then
        MsgBox "No client file selected!"
        Exit Sub
    End If
    
    If fileJ.Value = "" Then
        MsgBox "No Juyo file selected!"
        Exit Sub
    End If
    
    ' Stores the files names, for later porpures
    Range("D2").Value = fileJ.Value
    Range("C2").Value = fileC.Value

    ' Clears the last used segments, if any.
    sheetsBx.Clear
    
    ' Sets the variable for workbook Client File
    On Error Resume Next
    Set wb1 = Workbooks(Range("C2").Value)
    
    If err.Number <> 0 Then
        If err.Number = 9 Then
            MsgBox err.Number & " | " & err.Description & err1
            debug.print "1"
            err.Clear
            Exit Sub
        End If
    End If
    
    With wb1
        .Activate
        .Unprotect
    End With
    
    ' Shows all the sheets in the listbox, so users can select them
    For Each ws1 In wb1.Worksheets
        ws1.Visible = xlSheetVisible
        Me.sheetsBx.AddItem ws1.Name
    Next ws1
    
    '// TODO Let the userform expand itself
    '// TODO Load segments used in Juyo

    ' Here the process of getting the files from Juyo will start
    ' Later on the user can match his segments with the segments used in Juyo
    
    wb.Activate
    ws.Select

    'Sets the variable for workbook JUYO File
    On error resume Next
    set wb2 = workbooks(range("D2").value)

    If err.Number <> 0 Then
        If err.Number = 9 Then
            MsgBox err.Number & " | " & erl & " | " &err.Description & err1
            debug.print "2"
            err.Clear
            Exit Sub
        End If
    End If
    
    Set ws2 = wb2.Worksheets("Sheet0")

    wb2.Activate
    ws2.Select

    If cells(1,1).value <> "DATE" then
        MsgBox "Wrong file selected, Range A1 should be 'DATE'"
        Exit sub
    End IF

    lastColumn = Range("A1").End(xlToRight).Column

    iSegment = Application.WorksheetFunction.Transpose(Range(Cells(1, 1), Cells(1, lastColumn)))
    
    For iNum = 2 To UBound(iSegment) Step 2
        Debug.Print Left(iSegment(iNum, 1), Len(iSegment(iNum, 1)) - 3)
        Me.segmentJuyobx.AddItem Left(iSegment(iNum, 1), Len(iSegment(iNum, 1)) - 3)
    Next iNum

    ' Enables extra buttons
    CmdSheets.Enabled = True
    segmentbx.Enabled = True
    OptionButton1.Enabled = True
    OptionButton2.Enabled = True

    wb.Activate
    ws.Select

End Sub

Private Sub CmdSheets_click() 'Here the non-selected sheets will be deleted.

    Dim i As Integer, str As String, Value As String
    Dim a As Integer, b As Integer, item As Variant
    Dim lookup_value As String
    Dim fuzzyMonths() As Variant

    ' Select the correct workbook + sheets
    wb.Activate
    ws.Select
    
    Dim iSegments() As Variant
    Dim x As Integer, count As Integer, iNum As Integer
    
    ' This For Loop will delete all non-selected sheets.
    count = 1
    For x = 0 To sheetsBx.ListCount - 1
        If sheetsBx.Selected(x) = True Then
            ReDim Preserve iSegments(count)
            iSegments(count) = sheetsBx.List(x)
            count = count + 1
        End If
    Next x

    ' Exit sub if no sheets are selected.
    If count = 1 Then
        MsgBox "Select at least 1 sheet."
        exit sub
    End if

    ' Clears the whole listbox so it can be populated again
    sheetsBx.Clear
    
    'Dictonary for the months
    fuzzyMonths = Array("may", "May", "May", "may", _
                    "july", "July", "Jul", "jul", _
                    "june", "June", "Jun", "jun", _
                    "march", "March", "Mar", "mar", _
                    "april", "April", "Apr", "apr", _
                    "august", "August", "Aug", "Aug", _
                    "january", "January", "Jan", "jan", _
                    "october", "October", "Oct", "oct", _
                    "february", "February", "Feb", "feb", _
                    "december", "December", "Dec", "dec", _
                    "november", "November", "Nov", "nov", _               
                    "september", "September", "Sep", "sep", "Sept", "sept")
                    
                    
                    

    ' Here the selected segments that where selected will be added again
    ' Futhermore, with Fuzzy Seach, the months will be added.
    On Error Resume Next
    For iNum = 1 To UBound(iSegments)
        Me.sheetsBx.AddItem iSegments(iNum)

        lookup_value = iSegments(iNum)

        For Each item In fuzzyMonths
        
            str = item
            
            For i = 1 To Len(lookup_value)
                If InStr(item, Mid(lookup_value, i, 1)) > 0 Then
                    a = a + 1
                    item = Mid(item, 1, InStr(item, Mid(lookup_value, i, 1)) - 1) & Mid(item, InStr(item, Mid(lookup_value, i, 1)) + 1, 9999)
                End If                
            Next
            
            a = a - Len(item)
            
            If a > b Then
                b = a '1
                Value = str
            End If
            
            a = 0
            
        Next item
        
        Debug.Print Value
        
        Select Case Value
            Case "january", "January", "Jan", "jan"
                me.monthsBx.AddItem "january"
            Case "february", "February", "Feb", "feb"
                me.monthsBx.AddItem "february"
            Case "march", "March", "Mar", "mar"
                me.monthsBx.AddItem value
            Case "april", "April", "Apr", "apr"
                me.monthsBx.AddItem "march"
            Case "may", "May", "May", "may"
                me.monthsBx.AddItem "may"
            Case "june", "June", "Jun", "jun"
                me.monthsBx.AddItem "june"
            Case "july", "July", "Jul", "jul"
                me.monthsBx.AddItem "july"
            Case "august", "August", "Aug", "aug"
                me.monthsBx.AddItem "august"
            Case "september", "September", "Sep", "sep", "Sept", "sept"
                me.monthsBx.AddItem "september"
            Case "october", "October", "Oct", "oct"
                me.monthsBx.AddItem "october"
            Case "november", "November", "Nov", "nov"
                me.monthsBx.AddItem "november"
            Case "december", "December", "Dec", "dec"
                me.monthsBx.AddItem "december"
        End Select

        b = 0

    Next iNum
    
    Set ws1 = wb1.Worksheets(iSegments(1))

    if me.segmentJuyobx.ListCount <> me.segmentbx.ListCount then
        me.label13.ForeColor = RGB(255, 0, 0) ' red
    Else
        me.label13.Top =  1000
    End if

    CmdSegments.Enabled = True
    CmdLastUsedSeg.Enabled = True
    cmdTerminology.Enabled = True
    CmdStoreSegments.Enabled = True
End Sub

Private Sub CmdSegments_Click() 'Get the names of the segments.

    'Temporarily Hide Userform
    Me.Hide
    segmentbx.Clear

    wb1.Activate
    ws1.Select
    
    'Get Cell adress with values
    On Error Resume Next
        Set rng = Application.InputBox(Title:="Please select a range with all the segments names.", Prompt:="Select range of segments.", Type:=8)
    On Error GoTo 0

    If rng Is Nothing Then Exit Sub
    
    'Only let multiple selection through, otherwise it can be not wise.
    If rng.Cells.count = 1 Then
        MsgBox "You've selected only one cell." & "Please select multiple cells containing segments.", vbOKOnly
        Exit Sub
    End If

    'Get the values into a listbox for validation
    For Each cel In rng.Cells
        If cel.Value <> "" Then
            Debug.Print cel.Value
            Me.segmentbx.AddItem cel.Value
            Me.segmentSortbx.AddItem cel.Value
        End If
    Next

    set fullRange = rng
    
    wb.Activate
    ws.select

    'Unhide Userform
    Me.Show

End Sub  
'// TODO Find a way when there is only ADR and no REV
Private Sub cmdTerminology_Click() 'This will get the names of RN and REV

    'Temporarily Hide Userform
    Me.Hide
    terminologybx.Clear
    
    wb1.Activate
    ws1.Select
    
    'Get Cell adress with values
    On Error Resume Next
        Set rng = Application.InputBox(Title:="Please select a range with all the terminology names.", Prompt:="Select range of terminology.", Type:=8)
    On Error GoTo 0

    If rng Is Nothing Then Exit Sub
    
    'Only let multiple selection through, otherwise it can be not wise.
    If rng.Cells.count = 1 Then
        MsgBox "You've selected only one cell." & "Please select multiple cells containing terminology.", vbOKOnly
        Exit Sub
    End If

    'Get the values into a listbox for validation
    For Each cel In rng.Cells
        If cel.Value <> "" Then
            Debug.Print cel.Value
            Me.terminologybx.AddItem cel.Value
        End If
    Next

    set fullRange1 = rng

    wb.Activate
    ws.select
    
    'Unhide Userform
    Me.Show

End Sub

Private Sub CmdLastUsedSeg_Click() ' Will retreive last segments used

    Dim iMonth() as Variant
    Dim iNum as Integer
    Dim lastrow as Long

    segmentbx.Clear

    wb.Activate
    ws.select

    lastrow = Cells(Rows.count, "B").End(xlUp).Row

    iMonth = Range("B2:B" & lastrow)

    for iNum = 1 to UBound(iMonth)
        Me.segmentbx.AddItem iMonth(iNum, 1)
    Next iNum

End Sub

Private Sub CmdStoreSegments_Click() ' This will store the segments

    Dim x as Integer
    Dim lastrow as long

    wb.Activate
    ws.select

    lastrow = Cells(Rows.count, "B").End(xlUp).Row

    '//FIXME is nothing is loaded, the column will be deleted.
    Range("B2:B" & lastrow).ClearContents
    Range("B2").Select
    
    For x = 0 To Me.segmentbx.ListCount - 1
        Me.segmentbx.Selected(x) = True
        If Me.segmentbx.Selected(x) = True Then
            ActiveCell = Me.segmentbx.List(x)
            ActiveCell.Offset(1, 0).Select
        End If
    Next x

End sub

Private Sub segmentbx_Change()

    if me.segmentJuyobx.ListCount <> me.segmentbx.ListCount then
        me.label13.ForeColor = RGB(255, 0, 0) ' red
    Else
        me.label13.Top =  1000
    End if

End Sub

Private Sub CmdConvert_Click() ' This is the last sub before converting

    Dim msg as Variant
    dim answer as Integer

    msg = ("Segments are not evenly distributed! Make sure that the segments in both listboxes are 100% correct. " & _
            vbNewLine & vbNewLine & _
            "Segments JUYO count  : " & Me.segmentJuyobx.ListCount & vbNewLine & _
            "Segments Client count : " & Me.segmentbx.ListCount & _
            vbNewLine & vbNewLine & _
            "If you want to continue with uneven matching segments press 'Yes', otherwise press 'no'" & vbNewLine & _
            "NOTE: the segments that don't have a match will have no data here and in Launchpad.")
    
    If Me.segmentJuyobx.ListCount = Me.segmentbx.ListCount Then
        Debug.Print "COUNT Segments: " & Me.segmentJuyobx.ListCount & " | " & Me.segmentbx.ListCount
    Else
        answer = MsgBox (msg, vbCritical + vbYesNo + vbDefaultButton2, "Segments are Not correct!")
        if answer = vbyes then
            Debug.Print "COUNT Segments: " & Me.segmentJuyobx.ListCount & " | " & Me.segmentbx.ListCount
        Else
            msgbox "Please make sure all the segments have a 100% match"
            Exit Sub
        end if
    End If

    If CheckBox2.Value = True Then
        Call CmdStoreSegments_Click
    End If

    Call main

End Sub

Private sub main()

    'Temporarily Hide Userform
    Me.Hide

    Dim StartTime   As Double
    Dim SecondsElapsed As Double

    'Dim x as long
    Dim iData() As Variant
    Dim iSegments() As Variant
    Dim iMonths() As Variant
    Dim iDays() as Variant, iData1() as Variant
    Dim iTerm() As Variant
    Dim Loc As Range
    dim mDay as long, i as long, y as long
    Dim x As Integer, count As Integer
    Dim iNum As Integer, iNum1 As Integer, iNum2 As Integer
    Dim match_1 As Variant, match_2 As Variant, match_3 As Variant

    StartTime = Timer

    wb.Activate
    ws.Select

    iSegments() = segmentbx.List
    iMonths() = sheetsBx.List
    iTerm() = terminologybx.List
    iDays() = monthsBx.List

    With wb1
        .Activate
        .Unprotect
    End With

    'Here start the proces of populating the array
    For iNum = 0 To UBound(iMonths)

        On Error Resume Next
        Set ws1 = wb1.Worksheets(iMonths(iNum, 0))
    
        Debug.Print err.Number & " | " & err.Description

        If err.Number <> 0 Then
            If err.Number = 9 Then
                MsgBox "Month: '" & iMonths(iNum, 1) & "'. Is not recognised in as sheet in: " & wb1.Name & vbNewLine & _
                    vbNewLine & "Please try again or contact the admin.", vbCritical
                    err.clear
                Exit Sub
            End If
        End If

        ws1.Visible = xlSheetVisible
        ws1.Select

        Select Case iDays(iNum, 0)
            Case "january", "march", "may", "july", "august", "october", "december": mDay = 31
            Case "april", "june", "september", "november": mDay = 30
            Case "february": mDay = 28
            Case Else
                Debug.Print "No Month"
                mDay = 31 
        End Select

        With ws1.UsedRange

            Set Loc = .Cells.Find(What:=iTerm(0, 0)) 'Change to variable (iNum1,0)
            
            count = 0

            If Not Loc Is Nothing Then
                
                Do Until count = UBound(iSegments) + 1
                    
                    if OptionButton6 = true then
                    'Columns stored, so loc and fullrange1 must be row

                        If Loc.Row = fullRange1.Row Then
                            
                            Debug.Print Loc.Address
                            
                            ReDim Preserve iData(x)
                            iData(x) = Application.WorksheetFunction.Transpose(Range(Cells(loc.Row + 1, loc.Column), Cells(loc.Row + mDay, loc.Column)))
                                    
                            x = x + 1
                            
                            Set Loc = .FindNext(Loc)
                            count = count + 1
                        Else
                            Debug.Print Loc.Address
                            Set Loc = .FindNext(Loc)
                        End If
                    
                    Else

                        If Loc.Column = fullRange1.Column Then

                            Debug.Print Loc.Address

                            ReDim Preserve iData(x)
                            iData(x) = Application.WorksheetFunction.Transpose(Range(Cells(loc.Row, loc.Column +1), Cells(loc.Row, loc.Column + mDay)))
                            
                            x = x + 1

                            Set Loc = .FindNext(Loc)
                            count = count + 1
                        Else
                            Debug.Print Loc.Address
                            Set Loc = .FindNext(Loc)
                        End If

                    End if

                Loop
            
            End If
        
        End With
        
        Set Loc = Nothing
        x = 0

        With ws1.UsedRange

            Set Loc = .Cells.Find(What:=iTerm(1, 0)) 'Change to variable (iNum1,0)
            
            count = 0

            If Not Loc Is Nothing Then
                
                Do Until count = UBound(iSegments) + 1
                    
                    if OptionButton6 = true then
                    'Columns stored, so loc and fullrange1 must be row

                        If Loc.Row = fullRange1.Row Then
                            
                            Debug.Print Loc.Address
                            
                            ReDim Preserve iData1(x)
                            iData1(x) = Application.WorksheetFunction.Transpose(Range(Cells(loc.Row + 1, loc.Column), Cells(loc.Row + mDay, loc.Column)))
                                    
                            x = x + 1
                            
                            Set Loc = .FindNext(Loc)
                            count = count + 1
                        Else
                            Debug.Print Loc.Address
                            Set Loc = .FindNext(Loc)
                        End If
                    
                    Else

                        If Loc.Column = fullRange1.Column Then

                            Debug.Print Loc.Address

                            ReDim Preserve iData1(x)
                            iData1(x) = Application.WorksheetFunction.Transpose(Range(Cells(loc.Row, loc.Column +1), Cells(loc.Row, loc.Column + mDay)))
                            
                            x = x + 1

                            Set Loc = .FindNext(Loc)
                            count = count + 1
                        Else
                            Debug.Print Loc.Address
                            Set Loc = .FindNext(Loc)
                        End If

                    End if

                Loop
            
            End If
        
        End With

        Set Loc = Nothing
    
    Next iNum

    Dim strFind As Variant
    Dim strStored As Variant
    Dim arrTemp As Variant, arrTemp1 As Variant
    Dim arrNew As Variant, arrNew1 As Variant
    
    i = 0
    
    ' Start sorting array
    For i = 0 To Me.segmentbx.ListCount - 1
        strFind = Me.segmentbx.List(i)

        For y = 0 To Me.segmentSortbx.ListCount - 1
            
            strStored = Me.segmentSortbx.List(y)
            
            Debug.Print "L: " & i & " " & Me.segmentbx.List(i); " | R: " & y & " " & Me.segmentSortbx.List(y)

            If strFind = strStored Then
            
                If i = y Then
                
                    Debug.Print "-- Level & Name Match: " & strFind & " = " & i
                
                Else

                    Debug.Print "- Name Match: " & strFind & " : " & i & "-" & y
                    Debug.Print "- reordering... "
                    
                    arrTemp = iData(y)
                    arrNew = iData(i)
    
                    arrTemp1 = iData1(y)
                    arrNew1 = iData1(i)
    
                    iData(y) = arrNew
                    iData(i) = arrTemp
    
                    iData1(y) = arrNew1
                    iData1(i) = arrTemp1
                
                End If
            
            Else

                'Do nothing
            
            End If
        
        Next y
        
        y = 0

    Next i


End sub

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
    
    Dim x As Integer
    
    For x = 0 To Me.segmentbx.ListCount - 1
        on error resume next
        If Me.segmentbx.Selected(x) = True Then
            ListBox4.AddItem Me.segmentbx.List(x)
            segmentbx.RemoveItem x
        End If
    Next x
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
End Sub
    
Private Sub CmdUp_Click()
    
    Application.ScreenUpdating = False

    Dim curIndex As Integer, othIndex As Integer
    Dim curVal As Variant, othVal As Variant
    
    With Me.segmentbx
        
        If .ListIndex = -1 Then Exit Sub
        If .ListIndex = 0 Then Exit Sub
        
        curIndex = .ListIndex
        
        curVal = .List(curIndex)
        othIndex = curIndex - 1
        othVal = .List(othIndex)
        
        .List(othIndex) = curVal
        .List(curIndex) = othVal
        
        .Selected(othIndex) = True
        .ListIndex = othIndex
        .Selected(curIndex) = False
        
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub CmdDown_Click()
    
    Application.ScreenUpdating = False

    Dim curIndex As Integer, othIndex As Integer
    Dim curVal As Variant, othVal As Variant
    
    With Me.segmentbx
        
        If .ListIndex = -1 Then Exit Sub
        If .ListIndex = .ListCount - 1 Then Exit Sub
        
        curIndex = .ListIndex
        
        curVal = .List(curIndex)
        othIndex = curIndex + 1
        othVal = .List(othIndex)
        
        .List(othIndex) = curVal
        .List(curIndex) = othVal
    
        .Selected(othIndex) = True
        .ListIndex = othIndex
        .Selected(curIndex) = False
        
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub CmdRight_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call CmdRight_Click
End Sub

Private Sub CmdLeft_Click()
    
    Application.ScreenUpdating = False

    Dim itemIndex As Variant
    
    'Move selected items to the Left
    
    'Remove item from the left and place it on the right.
    'Loop through the items
    For itemIndex = ListBox4.ListCount - 1 To 0 Step -1
        
        'Check if an item was selected.
        If ListBox4.Selected(itemIndex) Then
            
            'Move selected item to the right.
            segmentbx.AddItem ListBox4.List(itemIndex)
            
            'Remove selected item from the left.
            ListBox4.RemoveItem itemIndex
            
        End If
        
    Next itemIndex
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub CmdLeft_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call CmdLeft_Click
End Sub
