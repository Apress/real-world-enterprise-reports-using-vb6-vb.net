Dim oRS As New ADODB.Recordset
Dim cSQL As String

cSQL = "SELECT * FROM Product"

With oRS
    Set .ActiveConnection = oConn
    .CursorLocation = adUseClient
    .CursorType = adOpenForwardOnly
    .Open cSQL
End With

Set oRS.ActiveConnection = Nothing

Do While Not oRS.EOF

    Debug.Print oRS("descr")

    oRS.MoveNext

Loop

oRS.Close
Set oRS = Nothing
