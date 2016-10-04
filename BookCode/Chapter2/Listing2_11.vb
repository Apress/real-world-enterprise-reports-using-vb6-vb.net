        Dim objDR As OleDbDataReader
        Dim objCommand As OleDbCommand
        Dim cSQL As String

        cSQL = "SELECT * " & _
               "FROM Source " & _
               "ORDER BY descr"

        objCommand = New OleDbCommand(cSQL, oConn)

        objDR = objCommand.ExecuteReader()

        While objDR.Read

            Console.WriteLine(objDR.GetInt32(0))
            Console.WriteLine(objDR.GetString(1))

      End While

      objDR.Close()
