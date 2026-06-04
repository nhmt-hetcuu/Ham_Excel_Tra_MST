Attribute VB_Name = "Module1"

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Dim jsonCache As Object

Function mst(taxCode As String, field As String) As String
    Dim http As Object
    Dim url As String
    Dim response As String, compactResponse As String
    Dim apiKey As String
    Dim searchStr As String
    Dim startPos As Long, endPos As Long
    Dim cleanCode As String
    
    On Error GoTo ErrorHandler
    
    If jsonCache Is Nothing Then
        Set jsonCache = CreateObject("Scripting.Dictionary")
    End If
    
    cleanCode = Trim(taxCode)
    If Left(cleanCode, 1) = "'" Then cleanCode = Mid(cleanCode, 2)
    cleanCode = Replace(cleanCode, "'", "")
    cleanCode = Replace(cleanCode, " ", "")
    cleanCode = Replace(cleanCode, vb160, "")
    
    If cleanCode = "" Then
        mst = ""
        Exit Function
    End If
    
    Select Case LCase(Trim(field))
        Case "ten": apiKey = "name"
        Case "diachi": apiKey = "address"
        Case "trangthai": apiKey = "status"
        Case "mst": apiKey = "taxID"
        Case "loaihinhdn": apiKey = "orgType"
        Case "quanly": apiKey = "taxDepartment"
        Case "ngaycapnhat": apiKey = "updatedAt"
        Case Else: apiKey = field
    End Select

    If jsonCache.Exists(cleanCode) Then
        response = jsonCache(cleanCode)
    Else
        Sleep 250
        
        Set http = CreateObject("MSXML2.ServerXMLHTTP")
        http.setTimeouts 5000, 5000, 5000, 5000
        
        url = "https://api.xinvoice.vn/gdt-api/tax-payer-records/" & cleanCode
        
        http.Open "GET", url, False
        http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        http.send
        
        response = http.responseText
        jsonCache.Add cleanCode, response
    End If
    
    compactResponse = Replace(response, " ", "")
    
    If InStr(1, compactResponse, """success"":false") > 0 Then
        mst = "Chua dang ky MST"
        Exit Function
    End If
    
    searchStr = """" & apiKey & """:"""
    startPos = InStr(1, response, searchStr)
    
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        endPos = InStr(startPos, response, """")
        
        If endPos > startPos Then
            mst = Mid(response, startPos, endPos - startPos)
            Exit Function
        End If
    End If
    
    mst = "Khong tim thay truong thong tin"
    Exit Function

ErrorHandler:
    If Err.Number = -2147012894 Then
        mst = "Loi: Mang cham hoac API qua tai"
    Else
        mst = "Loi VBA: " & Err.Description
    End If
End Function
