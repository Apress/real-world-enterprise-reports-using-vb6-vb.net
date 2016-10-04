cSQLCount = "SELECT COUNT(*) FROM Employees WHERE "
cSQL = "SELECT * FROM Employees WHERE "
cWhere = " LastName = " + Chr(34) + txtLastName  + Chr(34)
cSQL = cSQL + cWhere + " ORDER by LastName, FirstName "
cSQLCount = cSQLCount + cWhere
Set oRS = oConn.Execute cSQLCount
If oRS.Fields(0) > 1000 Then
   MsgBox "This query will return too many records. " & _
  "Please narrow the search criteria.", vbOKOnly, PROGNAME
   Screen.MousePointer = vbDefault
   Exit Sub
End If
'If no more than 1000 records will be returned, then go get the actual data
Set oRS = oConn.Execute cSQL
