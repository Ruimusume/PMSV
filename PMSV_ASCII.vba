Function ConvertUnicodeArrary(str As String, digit As Integer) As String
    Dim i As Long
    Dim byteArray() As Byte
    Dim hexResult As String
    
    ' 输入验证
    If str = "" Or digit < 1 Or digit > Len(str) Then
        ConvertUnicodeArrary = "0000"
        Exit Function
    End If
    
    ' 获取指定位置的字符
    Dim currentChar As String
    currentChar = Mid(str, digit, 1)
    
    ' 转换为Byte数组
    byteArray = currentChar
    
    ' 转换为十六进制字符串并自动补0
    If UBound(byteArray) = 1 Then
        ' 普通字符
        hexResult = UCase(Hex(byteArray(1)) & Right("00" & Hex(byteArray(0)), 2))
    Else
        ' 代理对字符
        hexResult = UCase(Hex(byteArray(3)) & Right("00" & Hex(byteArray(2)), 2) & _
                          Hex(byteArray(1)) & Right("00" & Hex(byteArray(0)), 2))
    End If
    
    ' 确保是4位（如果不足补0）
    If Len(hexResult) < 4 Then
        hexResult = String(4 - Len(hexResult), "0") & hexResult
    End If
    
    ConvertUnicodeArrary = hexResult
End Function
