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

Public fullRange As Range

Const err1 As Variant = vbNewLine & vbNewLine & _
                        "Workbook name is not the same. Please try again."

' Ctrl + '/' to quicly comment

Private Sub UserForm_Initialize()

    Dim vWorkbook As Workbook
    
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

    Dim iSegment() As Variant
    Dim lastColumn As Long
    Dim iNum As Integer

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

    wb.Activate
    ws.Select
    
    Dim iSegments() As Variant
    Dim i As Integer, count As Integer, iNum As Integer
    
    count = 1
    For i = 0 To sheetsBx.ListCount - 1
        If sheetsBx.Selected(i) = True Then
            ReDim Preserve iSegments(count)
            iSegments(count) = sheetsBx.List(i)
            count = count + 1
        End If
    Next i

    If count = 1 Then
        MsgBox "Select at least 1 sheet."
        exit sub
    End if

    sheetsBx.Clear
    
    On Error Resume Next
    For iNum = 1 To UBound(iSegments)
        Me.sheetsBx.AddItem iSegments(iNum)
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
        End If
    Next
    
    wb.Activate
    ws.select

    'Unhide Userform
    Me.Show

End Sub  

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

Private Sub CmdConvert_Click() ' Here will be the final space before converting the numbers

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

    ' Start converting here.
    Call main

End Sub

Private sub main()

    'Temporarily Hide Userform
    Me.Hide

    Dim StartTime   As Double
    Dim SecondsElapsed As Double

    StartTime = Timer

    '// TODO start converting program

    wb.activate
    ws.select



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
