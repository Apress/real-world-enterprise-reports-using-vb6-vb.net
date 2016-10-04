Private Sub Page_Load(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles MyBase.Load

    Dim oConn As New OleDb.OleDbConnection()
    Dim objDR As OleDb.OleDbDataReader
    Dim cConnectString As String

    'Connect to data source
    cConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=C:\\BookCode\\SampleDatabase.mdb"

    oConn.ConnectionString = cConnectString
    oConn.Open()

    'Invoke routines which display web controls
    ShowDateRange(SELECT_ENTER_DATE, 8, 32, 20, "Enter Date:")
    ShowCheckBox(SELECT_COMPLEX_QUESTION, 8, 95, 20, "Complex Question")
    ShowComboBox(SELECT_DEPARTMENT, 8, 140, 180, 20, "Department")
    ShowListBox(SELECT_PRODUCT, 8, 200, 180, 20, "Product")
    ShowListBox(SELECT_SOURCE, 200, 200, 180, 20, "Source")

    'Populate data bound web controls (list and combo boxes) with data
    If Not Page.IsPostBack Then
        LoadTable(oConn, objComboBoxColl(SELECT_DEPARTMENT), _
            "Department", "ID", "Descr")
        LoadTable(oConn, objListBoxColl(SELECT_PRODUCT), _
            "Product", "ID", "Descr")
        LoadTable(oConn, objListBoxColl(SELECT_SOURCE), _
            "Source", "ID", "Descr")
    End If

    PositionButtons()

End Sub 'Page_Load
