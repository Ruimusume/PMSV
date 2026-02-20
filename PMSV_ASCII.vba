Function ConvertUnicodeArrary(str As String, digit As Integer) As String
    Dim i As Long
    Dim byteArray() As Byte
    Dim hexResult As String
    Dim charIndex As Integer
    
    ' 输入验证
    If str = "" Or digit < 1 Or digit > Len(str) Then
        ConvertUnicodeArrary = "0000"
        Exit Function
    End If
    
    ' 获取指定位置的字符
    Dim currentChar As String
    currentChar = Mid(str, digit, 1)
    
    ' 转换为Byte数组（VBA中String赋值给Byte数组会自动使用Unicode UTF-16LE编码）
    byteArray = currentChar
    
    ' 转换为十六进制字符串（注意：Byte数组在VBA中是从0开始的）
    hexResult = ""
    
    ' UTF-16LE编码：低位字节在前，高位字节在后
    ' 所以byteArray(0)是低位，byteArray(1)是高位
    For i = 0 To UBound(byteArray)
        ' 格式化为2位十六进制
        hexResult = hexResult & Right("00" & Hex(byteArray(i)), 2)
    Next i
    
    ' 确保是4位十六进制（2字节）
    If Len(hexResult) < 4 Then
        hexResult = hexResult & String(4 - Len(hexResult), "0")
    End If
    
    ' 返回结果（已经是低位在前，高位在后的格式）
    ConvertUnicodeArrary = UCase(hexResult)
    
End Function
