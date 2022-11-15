VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} init 
   Caption         =   "UserForm2"
   ClientHeight    =   12870
   ClientLeft      =   420
   ClientTop       =   1065
   ClientWidth     =   7530
   OleObjectBlob   =   "init.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "init"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public wb As Workbook, wb1 As Workbook
Public ws As Worksheet, ws1 As Worksheet

Public fullRange As Range

Const err1 As Variant = vbNewLine & vbNewLine & _
                        "Workbook name is not the same. Please try again."

' Ctrl + '/' to quicly comment
' 

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
    fileJ.ListIndex = 0
    
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
    Range("C2").Value = fileJ.Value
    Range("D2").Value = fileC.Value

    ' Clears the last used segments, if any.
    sheetsBx.Clear
    
    'Sets the variable for workbook JUYO File
    On Error Resume Next
    Set wb1 = Workbooks(Range("C2").Value)
    
    If err.Number <> 0 Then
        If err.Number = 9 Then
            MsgBox err.Number & " | " & err.Description & err1
            err.Clear
            Exit Sub
        End If
    End If
    
    With wb1
        .Activate
        .Unprotect
    End With
    
    ' Shows all the sheets in the listbox, so the users can select them
    For Each ws1 In wb1.Worksheets
        ws1.Visible = xlSheetVisible
        Me.sheetsBx.AddItem ws1.Name
    Next ws1
    
    wb.Activate
    ws.Select
    
    '// TODO Let the userform expand itself
    
    ' Enables extra buttons
    CmdSheets.Enabled = True
    segmentbx.Enabled = True
    OptionButton1.Enabled = True
    OptionButton2.Enabled = True

    Application.ScreenUpdating = True

End Sub

Private Sub CmdSheets_Click() 'Here the non-selected sheets will be deleted.

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
    
    Application.ScreenUpdating = True
    
    Call GetRangeSegments

End Sub

Public Sub GetRangeSegments() 'Here the names of the segments will be selected

    Dim rng As Range
    Dim cel As Range

    'Temporarily Hide Userform
    Me.Hide
    segmentbx.Clear
    
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
    
    'Unhide Userform
    Me.Show

End Sub

Private Sub CommandButton1_Click()
    Dim rng As Range
    Dim cel As Range

    'Temporarily Hide Userform
    Me.Hide
    
    'Get Cell adress with values
    On Error Resume Next
        Set rng = Application.InputBox(Title:="Please select a range", Prompt:="Select range", Type:=8)
    On Error GoTo 0

    If rng Is Nothing Then Exit Sub
    
    'Only let multiple selection through, otherwise it can be not wise.
    If rng.Cells.count = 1 Then
        MsgBox "You�ve selected only one cell." & "Please select multiple cells.", vbOKOnly
        Exit Sub
    End If

    'Get the values into a listbox for validation
    Set fullRange = rng
    
    'Unhide Userform
    Me.Show
End Sub