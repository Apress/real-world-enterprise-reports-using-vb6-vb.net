.Write("<html><body>")
.Write("<table>")
		
cCustomerID = oRS("CustomerID")
		
.Write("<TR><TD><A NAME=" & oRS("CustomerID") & "></A></TD></TR>")
.Write("<tr><td align=left colspan=2><B>" & _
	oRS("CompanyName") & "</B></td></tr>")

Do while not oRS.EOF

  .Write("</tr>")	
  .Write("<td align=left>" & oRS("OrderDate") & "</td>")
  .Write("<td align=right>" & oRS("Freight") & "</td>")
  .Write("</tr>")
		  
  dblTotal = dblTotal + CDbl((0 & oRS("Freight")))
  dblGrandTotal = dblGrandTotal + CDbl((0 & oRS("Freight")))
		  
  oRS.movenext
		  
  If oRS.EOF Then
	.Write("<tr height=10></tr>")
	.Write("<td align=right colspan=2>" & dblTotal & "</td>")
	Exit Do
  End if
		  
  If cCustomerID <> oRS("CustomerID") Then
			
	 cCustomerID = oRS("CustomerID")
				  
	.Write("<tr height=10></tr>")
	.Write("<td align=right  colspan=2>" & dblTotal & "</td>")
			
	 dblTotal = 0
			 		
	.Write("<tr height=10></tr>")
	.Write("<TR><TD><A NAME=" & oRS("CustomerID") & "></A></TD></TR>")		
	.Write("<tr><td align=left colspan=2><B>" & _
		oRS("CompanyName") & "</B></td></tr>")
		  
  End If

loop

.Write("<tr>")
.Write("</tr>")

.Write("<tr>")
.Write("<td><b>Grand Total</b></td>" )
.Write("<td align=right><b>" & dblGrandTotal & "</b></td>" )
.Write("</tr>")
				  	  
.Write("</table></body></html>")
