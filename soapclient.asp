<%
Class SoapClient

    '--- private
    Function Base64Encode(inData)
        Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
        Dim sOut, I
        For I = 1 To Len(inData) Step 3
            Dim nGroup, pOut
            nGroup = &H10000 * Asc(Mid(inData, I, 1)) + _
            &H100 * ZeroASC(Mid(inData, I + 1, 1)) + _
            ZeroASC(Mid(inData, I + 2, 1))
            nGroup = Oct(nGroup)
            nGroup = String(8 - Len(nGroup), "0") & nGroup
            pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
            Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
            Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
            Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
            sOut = sOut + pOut    
        Next
        Select Case Len(inData) Mod 3
            Case 1: '8 bit final
            sOut = Left(sOut, Len(sOut) - 2) + "=="
            Case 2: '16 bit final
            sOut = Left(sOut, Len(sOut) - 1) + "="
        End Select
        Base64Encode = sOut
    End Function

    Function ZeroASC(OneChar)
        If OneChar = "" Then ZeroASC = 0 Else ZeroASC = Asc(OneChar)
    End Function
    '--- end private

    '--- değişkenler
    Public ObjXMLHttp
    '--- işlemlere başlamadan önce bu iki değişkeni doldurmak gerekir
    Public Username
    Public Password

    '-- data gönderip string sonucu döndürüp işini bitirir. ObjXMLHTTP yıkılır
    Public Function SendData(ServiceUrl, Data)
        Set ObjXMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
        ObjXMLHttp.SetTimeouts 10000, 60000 , 60000, 360000 
        ObjXMLHttp.Open "POST", ServiceUrl, False
        If Not (IsEmpty(Username) Or IsNull(Username) Or Username = "") Then 
            ObjXMLHttp.SetRequestHeader "Authorization", "Basic " & Base64Encode(Username&":"&Password)
        End If
        ObjXMLHttp.SetRequestHeader "Content-Type", "text/xml; charset=utf-8"
        ObjXMLHttp.SetRequestHeader "Content-Length", Len(Data)
        ObjXMLHttp.Send Data
        SendData = ObjXMLHTTP.ResponseText
        Call Dispose
    End Function

    '-- data gönderip nesneyi kullanmak için açık bırakır. İş bitiminde Dispose çağrılmalıdır.
    '-- ObjXMLHttp den veriyi responsexml veya responsestream olarak okumak için kullanabilirsiniz
    Public Function SendDataAndWait(ServiceUrl, Data)
        Set ObjXMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
        ObjXMLHttp.SetTimeouts 10000, 60000 , 60000, 360000 
        ObjXMLHttp.Open "POST", ServiceUrl, False
        If Not (IsEmpty(Username) Or IsNull(Username) Or Username = "") Then 
            ObjXMLHttp.SetRequestHeader "Authorization", "Basic " & Base64Encode(Username&":"&Password)
        End If
        ObjXMLHttp.SetRequestHeader "Content-Type", "text/xml; charset=utf-8"
        ObjXMLHttp.SetRequestHeader "Content-Length", Len(Data)
        ObjXMLHttp.Send Data
    End Function

    '-- İşlemler bittikten sonra nesneyi yıkmaya yarar
    Public Sub Dispose()
        Set ObjXMLHTTP = Nothing
    End Sub

End Class
%>