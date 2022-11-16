Attribute VB_Name = "Devxlwings"
Sub FuzzySeach()
    Dim i As Integer, str As String, Value As String
    Dim a As Integer, b As Integer, item As Variant
    Dim lookup_value As String
    Dim fuzzyMonths() As Variant
    
    
    
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

    lookup_value = "DDecember 2022"

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
            Debug.Print "january"
        Case "february", "February", "Feb", "feb"
            Debug.Print "february"
        Case "march", "March", "Mar", "mar"
            Debug.Print "march"
        Case "april", "April", "Apr", "apr"
            Debug.Print "april"
        Case "may", "May", "May", "may"
            Debug.Print "may"
        Case "june", "June", "Jun", "jun"
            Debug.Print "june"
        Case "july", "July", "Jul", "jul"
            Debug.Print "july"
        Case "august", "August", "Aug", "aug"
            Debug.Print "august"
        Case "september", "September", "Sep", "sep", "Sept", "sept"
            Debug.Print "september"
        Case "october", "October", "Oct", "oct"
            Debug.Print "october"
        Case "november", "November", "Nov", "nov"
            Debug.Print "november"
        Case "december", "December", "Dec", "dec"
            Debug.Print "december"
    End Select

    b = 0

End Sub

