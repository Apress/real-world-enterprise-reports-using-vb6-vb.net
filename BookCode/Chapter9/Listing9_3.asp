Sub CreateListBox(oConn, cCaption, cName, cTable, cID, cDescr, iTop, iLeft)
	Dim oRS
	Dim cSQL 

	cSQL = "SELECT " & cID & ", " & cDescr & _
		   " FROM " & cTable & _
		   " ORDER BY " & cDescr

	Set oRS = oConn.Execute(cSQL) 

	With Response
		.Write("<DIV id=mydiv style='position:absolute; top:" & _
                   iTop & "px; left:" & iLeft & "px;'>")

		.Write("<TABLE>")
	
		.Write("<TR>")
		.Write("<TD>")

		.Write(cCaption) 

		.Write("</TD>")
		.Write("</TR>")

		.Write("<TR>")
		.Write("<TD>")
	
		.Write("<select name=" & cName & " multiple  style=height:120px>") 

		Do While not oRS.EOF 
		
			.Write("<option value=" & chr(39) & oRS(cID) & _
                         "," & oRS(cDescr) & chr(39) & ">" & oRS(cDescr))

			oRS.MoveNext 

		Loop

		.Write("</select>")

		.Write("</TD>")
		.Write("</TR>")

		.Write("<TR>")
		.Write("<TD>")

		.Write("<INPUT type=button value='Clear all '" & cCaption & _
                   " id=button1 name=button1 onclick=cmdClearAll_Click('" & _
                   cName & "')>") 

		.Write("</TD>")
		.Write("</TR>")

		.Write("</TABLE>")
		.Write("</DIV>")		
	End With
	
	oRS.Close
	Set oRS = Nothing
		
End Sub
