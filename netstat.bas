Attribute VB_Name = "Module1"
Public Sub AddData()
Dim x As Integer
Dim strColumn As String
Dim strIPC As String
Dim strIsp As String
Dim strOrg As String
Dim strTest As String


Dim strIP As String


x = 2
strColumn = "A"
strIPC = "G"
strIsp = "L"
strOrg = "M"

Do While Not IsEmpty(Range(strColumn + CStr(x)).Value)

    strIP = Range(strIPC + CStr(x)).Value
    
    If IsPublic(strIP) Then
       strTest = GetApiData(strIP, "isp")
       Range(strIsp + CStr(x)).Value = strTest
       strTest = GetApiData(strIP, "org")
       Range(strOrg + CStr(x)).Value = strTest
    End If
    
    x = x + 1

Loop

End Sub

Function GetApiData(strIP As String, strType As String) As String
Dim XMLHTTP As Object
Dim strResponse As String
Dim strURL As String
Dim intStart As Integer
Dim intLen As Integer

    strURL = "http://ip-api.com/json/" + strIP
    Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")
    XMLHTTP.Open "GET", strURL, False
    XMLHTTP.send
    strResponse = XMLHTTP.responsetext
    'strResponse = Replace(strResponse, Chr(10), vbNewLine)
    
    Debug.Print strResponse
    Set XMLHTTP = Nothing
     
    intStart = InStr(strResponse, strType) + 6
    intLen = InStr(intStart, strResponse, """,") - intStart
    Debug.Print Mid(strResponse, intStart, intLen)
    
    
    GetApiData = Mid(strResponse, intStart, intLen)
    
   
End Function

Function IsPublic(strIP As String) As Boolean
Dim strFirst3 As String
Dim strFirst As String
Dim strSecond As String
Dim y As Integer

strFirst3 = Left(strIP, 3)
strFirst = Left(strIP, 1)

If strFirst3 = "172" Then
    strSecond = Mid(strIP, 5, InStr(5, strIP, ".") - 4)
    y = CInt(strSecond)
    
    If y < 16 Or y > 31 Then
        IsPublic = True
    Else
        IsPublic = False
    End If
ElseIf strFirst3 <> "127" And strFirst <> ":" And strFirst <> "0" And strFirst3 <> "192" Then
    IsPublic = True
Else
    IsPublic = False
End If



End Function
