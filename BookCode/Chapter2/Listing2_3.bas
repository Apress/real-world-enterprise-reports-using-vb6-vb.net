Dim oCmd As New ADODB.Command
Dim oRS As New ADODB.Recordset
Dim cSQL As String

cSQL = "SELECT * FROM Product"

With oCmd
    Set .ActiveConnection = oConn
    .CommandText = cSQL
    .CommandType = adCmdText
    .CommandTimeout = 20
    Set oRS = .Execute
End With
