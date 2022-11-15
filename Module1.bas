Attribute VB_Name = "Module1"
Const wsDataName As Variant = "Sheet1"
Sub test123()

err.Clear

End Sub

Sub SelRange()


Dim oFound As Range
Dim oLookin As Range
Dim sLookFor As String

sLookFor = "July" 'Change to suit

Set oLookin = Worksheets("DJuly 2022").UsedRange 'Change sheet name to suit

Set oFound = oLookin.Find(what:=sLookFor, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)

If Not oFound Is Nothing Then
    MsgBox oLookin.Address
End If

End Sub


    Set wb = Workbooks(Range("E2").Value)
    
    If err.Number <> 0 Then
        If err.Number = 9 Then
            Debug.Print err.Number & " | " & err.Description
            Set wb = Workbooks(Range("E2").Value & ".xlsm")
            err.Clear
        End If
    End If
