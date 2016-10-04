        Dim objDR As OleDbDataReader
        Dim objCommand As New OleDbCommand()
        Dim cSQL As String

        cSQL = "Source"

        With objCommand
            .Connection = oConn
            .CommandText = cSQL
            .CommandType = CommandType.TableDirect
            .CommandTimeout = 60
            objDR = .ExecuteReader()
      End With
