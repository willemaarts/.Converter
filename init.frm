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

    Application.ScreenUpdating = False
    
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
    
    ' This will disable, so that the users can't click anything else.
    ' Now everything is in the right order
    CmdLoad.Enabled = True
    CmdSheets.Enabled = False
    segmentbx.Enabled = False
    OptionButton1.Enabled = False
    OptionButton2.Enabled = False
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub CmdLoad_Click() 'This will start the first process.

    Application.ScreenUpdating = False

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

    Application.ScreenUpdating = True

End Sub

Private Sub CmdSheets_click() 'Here the non-selected sheets will be deleted.

    Application.ScreenUpdating = False
    
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

    sheetsBx.Clear
    
    On Error Resume Next
    For iNum = 1 To UBound(iSegments)
        Me.sheetsBx.AddItem iSegments(iNum)
    Next iNum
    
    Set ws1 = wb1.Worksheets(iSegments(1))
    
    Application.ScreenUpdating = True
    
    'Call GetRangeSegments

End Sub

Private Sub CmdSegments_Click() 'Get the names of the segments.

    'Temporarily Hide Userform
    Me.Hide
    segmentbx.Clear

    Application.ScreenUpdating = True
    
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

Private Sub cmdTerminology_Click()

    'Temporarily Hide Userform
    Me.Hide
    terminologybx.Clear

    Application.ScreenUpdating = True
    
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
        
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub CmdDown_Click()
    
    Application.ScreenUpdating = False
    
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
