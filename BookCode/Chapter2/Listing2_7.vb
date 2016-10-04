        Dim objDR As OleDbDataReader
        Dim objCommand As New OleDbCommand()
        Dim cSQL As String

        cSQL = "SELECT * FROM Source"

        With objCommand
            .Connection = oConn
            .CommandText = cSQL
            .CommandType = CommandType.Text
            .CommandTimeout = 60
            objDR = .ExecuteReader()
       End With
