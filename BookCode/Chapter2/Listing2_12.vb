cSQL =“SELECT COUNT(*)“&_
      “FROM Source ”

objCommand =New OleDbCommand(cSQL,oConn)

Console.WriteLine(objCommand.ExecuteScalar)