Private Sub BindToDataSource()
    Dim oConn As New OleDb.OleDbConnection()
    Dim objDA As New OleDb.OleDbDataAdapter()
    Dim objCommand As New OleDb.OleDbCommand()
    Dim objDS As New DataSet()
    Dim objDV As New DataView()
    Dim cConnectString As String
    Dim cSQL As String

    cConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=C:\\BookCode\\SampleDatabase.mdb"

    oConn.ConnectionString = cConnectString
    oConn.Open()

    cSQL = "SELECT ID, ConventionName, StartDate, City, State " & _
           "FROM Convention"

    With objCommand
        .Connection = oConn
        .CommandText = cSQL
    End With

    objDA.SelectCommand = objCommand
    objDA.Fill(objDS, "Convention")

    objDV.Table = objDS.Tables("Convention")
    objDV.Sort = ViewState("SortOrder")

    DataGrid1.DataSource = objDV
    DataGrid1.DataBind()
End Sub
