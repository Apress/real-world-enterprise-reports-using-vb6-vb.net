cSQL = "SELECT DISTINCT c.CustomerID, c.CompanyName " & _
          "FROM Customers c, Orders o " & _
          "WHERE c.CustomerID = o.CustomerID " & _ 
          "ORDER BY c.CompanyName"

Set oRS = oConn.Execute(cSQL)

Response.Write("<TABLE>")
     
Do While Not oRS.EOF     

     Response.Write("<TR><TD><A HREF=#" & oRS("CustomerID") & ">" & _
          oRS("CompanyName") & "</A></TD></TR>")

     oRS.MoveNext 
Loop

Response.Write("</TABLE>")
