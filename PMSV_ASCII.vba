Attribute VB_Name = "Ä£¿é1"
Function ConvertUnicodeArrary(str As String, digit As Integer) As String
    Dim arr() As Byte
    Dim i
    Dim char() As String
    If str = "" Or digit > Len(str) Then
        ConvertUnicodeArrary = "0"
        Exit Function
    End If
    
    ReDim char(Len(str))
    
    For x = 1 To Len(str)
        char(x - 1) = Mid(str, x, 1)
    Next x
    For i = 0 To UBound(char)
        ReDim arr(0 To UBound(char) * 2)
        arr = char(i)
         char(i) = ""
        For j = 0 To UBound(arr)
         If IsNumeric(VBA.Hex(arr(j))) And Len(VBA.Hex(arr(j))) = 1 Then
            char(i) = "0" & VBA.Hex(arr(j)) & char(i)
         Else
            char(i) = VBA.Hex(arr(j)) & char(i)
        End If
        
        Next j
    Next i
    ConvertUnicodeArrary = char(digit - 1)
    
End Function
