<WebMethod()> Public Function SendReportToUser(ByVal cFileName As String) As String

    SendReportToUser = Base64Encode(cFileName)

End Function

Private Function Base64Encode(ByVal cFileName As String) As String
    Dim objFileStream As System.IO.FileStream
    Dim aData() As Byte
    Dim lBytes As Long
    Dim cResult As String

    objFileStream = New System.IO.FileStream(cFileName, _
        IO.FileMode.Open, IO.FileAccess.Read)

    ReDim aData(objFileStream.Length)

    lBytes = objFileStream.Read(aData, 0, objFileStream.Length)

    objFileStream.Close()

    cResult = System.Convert.ToBase64String(aData, 0, aData.Length)

    Base64Encode = cResult

End Function
