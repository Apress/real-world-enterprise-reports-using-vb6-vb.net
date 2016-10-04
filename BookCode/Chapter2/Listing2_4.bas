Dim oCmd As New ADODB.Command
Dim oRS As New ADODB.Recordset
Dim cSQL As String

cSQL = "sp_GetProduct"

With oCmd
    Set .ActiveConnection = oConn
    .CommandText = cSQL
    .CommandType = adCmdStoredProc
    .CommandTimeout = 20
    .Parameters.Append .CreateParameter("ID", adInteger, adParamInput, , 1)
    Set oRS = .Execute
End With
