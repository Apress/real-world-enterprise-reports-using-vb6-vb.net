<html>
<title>Customer Freight Report</title>
<body bgcolor="#ffffff">

</body>
</html>


<%
Dim oConn
Dim oRS
Dim cnstr
Dim cSQL
Dim dblTotal

set oConn=server.createobject("ADODB.connection")
set oRS = Server.CreateObject("ADODB.Recordset")

cConnectString = "Provider=SQLOLEDB.1;" & _
		 "Password=sa;" & _
		 "Persist Security Info=True;" & _
		 "User ID=testuser;" & _
		 "Initial Catalog=NorthWind;" & _
		 "Data Source=SETON-LAPTOP"

oConn.open cConnectString

cSQL = "SELECT c.CustomerID, c.CompanyName, c.ContactName, " & _
		"SUM(o.Freight) AS TotalFreight " & _
		"FROM Customers c, Orders o " & _
		"WHERE c.CustomerID = o.CustomerID " & _
		"GROUP BY c.CustomerID, c.CompanyName, c.ContactName  " & _
		"ORDER BY c.CompanyName"

Set oRS = oConn.Execute(cSQL)

With Response
	.Write("<font>")
	
	.Write("<table>")
	.Write("<tr>")
	.Write("<td width = ""350"" nowrap align=""left"">Seton Software</td>")
	.Write("<td width = ""140"" nowrap align=""right"">" & DATE & "</td>")	
	.Write("</tr>")
	.Write("<tr>")
	.Write("</tr>")		
	.Write("</table>")
	
	.Write("<table>")
	.Write("<tr>")
	.Write("<th width = ""500"" nowrap align=""center"">Freight Report </th>")
	.Write("</tr>")	
	.Write("</table>")
	
	
	.Write("<html><body>")
	.Write("<table>")
	.Write("<tr>")
	.Write("<th nowrap align=""left"">Company Name</th>")
	.Write("<th nowrap align=""left"">Contact Name</th>")
	.Write("<th nowrap align=""right"">Total Freight</th>")
	.Write("</tr>")

	Do while not oRS.EOF
	  .Write("<tr>")
	  .Write("<td align=""left""><a href=customerdetails.asp?id=" & _
	    oRS("CustomerID") & ">" & oRS("CompanyName") & "</a></td>" )
	  .Write("<td align=""left"">" & oRS("ContactName") & "</td>" )
	  .Write("<td align=""right"">" & FormatCurrency(0 & _
	    oRS("TotalFreight"),2) & "</td>" )
	  .Write("</tr>")
	  
	  dblTotal = dblTotal + CDbl((0 & oRS("TotalFreight")))
	  
	  oRS.movenext

	loop

	.Write("<tr>")
	.Write("</tr>")

	.Write("<tr>")
	.Write("<td></td>" )
	.Write("<td><b>Grand Total</b></td>" )
	.Write("<td align=""right""><b>" & dblTotal & "</b></td>" )
	.Write("</tr>")
			  	  
	.Write("</table></body></html>")

End With

oRS.Close
oConn.Close

Set oRS = nothing
Set oConn = nothing
Set oCmd = nothing

%>
