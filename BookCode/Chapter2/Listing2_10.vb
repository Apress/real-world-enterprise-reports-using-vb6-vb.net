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
            .Prepare()
            objDR = .ExecuteReader()
        End With

        Console.WriteLine(objDR.Item("Descr"))

        objDR.Close()

        With objCommand
            .Connection = oConn
            .CommandText = cSQL
            .CommandType = CommandType.StoredProcedure
            .CommandTimeout = 60
            .Parameters(0).Value = 2
            .Prepare()
            objDR = .ExecuteReader()
        End With

     Console.WriteLine(objDR.Item("Descr"))
