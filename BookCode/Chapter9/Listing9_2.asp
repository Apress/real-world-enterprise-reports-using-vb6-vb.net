<%

Dim oConn
Dim iReport

iReport = Request.Form("Report") 

Set oConn = Server.CreateObject("ADODB.Connection") 

oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=C:\Docs\Reports\SampleDatabase.mdb;" & _
"Persist Security Info=False"
oConn.Open

Select Case iReport

	Case 0
		Call CreateListBox(oConn, "Products", "lstProduct", _
			"Product", "ID", "Descr", 60, 120)
		Call CreateListBox(oConn, "Department", "lstDepartment", _
			"Department", "ID", "Descr", 60, 320)
		Call CreateComboBox(oConn, "Source:", "cboSource", _
			"Source", "ID", "Descr", 280, 120)
		Call CreateCheckBox("Closed Orders only:", "cboClosed", _
			"Ignore This", 320, 120)
		Call CreateDateRange("Range of Ship Dates:", _
			"txtShip", 380, 120 )
		Call CreateDateRange("Range of Order Dates:", _
			"txtOrder", 420, 120)
		Call DisplayButtons(480, 300)
		
End Select

%>

<INPUT type="hidden"  id=text1 name=txtReport value=<%=iReport%>>
