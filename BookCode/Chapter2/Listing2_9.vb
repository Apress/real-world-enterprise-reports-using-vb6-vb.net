        Dim objDR As OleDbDataReader
        Dim objCommand As New OleDbCommand()
        Dim cSQL As String

        cSQL = "sp_GetProduct"

        With objCommand
            .Connection = oConn
            .CommandText = cSQL
            .CommandType = CommandType.StoredProcedure
            .CommandTimeout = 60
            .Parameters.Add("ProductID", OleDbType.Integer).Value = 1
            objDR = .ExecuteReader()
       End With
