<%
Dim objReport
Dim cFileName
Dim cURL

Response.Buffer = "true"

Set objReport = CreateObject("Reports.Report")

With objReport
    .Destination = 0
    .UserID = 12
    .IsWebReport = True
    cFileName = .MyFirstWebReport
End With

Set objReport = Nothing

cURL = "http://myserver/" & cFileName

Response.Write("<script>" & vbCrLf) 
Response.Write("window.open('" & cURL & "');" & vbCrLf) 
    Response.Write("</script>")

%>
