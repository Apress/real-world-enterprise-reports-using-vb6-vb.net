<%@ Language=VBScript %>
<%

basePath = Request.ServerVariables("PATH_TRANSLATED")
While (Right(basePath, 1) <> "\" And Len(basePath) <> 0)
    iLen = Len(basePath) - 1
    basePath = Left(basePath, iLen)
Wend

baseVirtualPath = Request.ServerVariables("PATH_INFO")
While (Right(baseVirtualPath, 1) <> "/" And Len(baseVirtualPath) <> 0)
    iLen = Len(baseVirtualPath) - 1
    baseVirtualPath = Left(baseVirtualPath, iLen)
Wend

If Not IsObject(session("oApp")) Then
    Set session("oApp") = Server.CreateObject("CrystalRuntime.Application")
    If Not IsObject(session("oApp")) Then
        response.write "Error:  Could not instantiate the Crystal Reports... "
        response.end
    End If
End If   

If IsObject(session("oRpt")) then
     set session("oRpt") = nothing
End If

reportFileName = "Orders.rpt"

Set session("oRpt") = session("oApp").OpenReport(basepath & reportFileName, 1)
If Err.Number <> 0 Then
  Response.Write "Error Occurred creating Report Object: " & Err.Description
  Set Session("oRpt") = Nothing
  Set Session("oApp") = Nothing
  Session.Abandon
  Response.End
End If

if not session("oRpt").HasSavedData then
    Response.Write "<strong>ReCrystallize</strong><br><br>You are...</b><br>"
    Response.End
end if

session("oRpt").MorePrintEngineErrorMessages = False
session("oRpt").EnableParameterPrompting = False

If IsObject (session("oPageEngine")) Then
   set session("oPageEngine") = nothing
End If

set session("oPageEngine") = session("oRpt").PageEngine
%>
<% viewer = "PDF" %>
<%
if viewer = "PDF" then
    exporttype = "31"
    fileextension = ".pdf"
end if
%>
<%
set crystalExportOptions = Session("oRpt").ExportOptions
ExportFileName = "Orders report-" & CStr(Session.SessionID) & fileextension
ExportDirectory = basePath

crystalExportOptions.DiskFileName = basepath & ExportFileName
crystalExportOptions.FormatType = CInt(exporttype)
crystalExportOptions.DestinationType = 1
Session("oRpt").Export False

Set Session("oPageEngine") = Nothing
Set Session("oRpt") = Nothing
'Set Session("oApp") = Nothing

response.write "<META http-equiv=" & Chr(34) & "Refresh" & Chr(34) & _
" content=" & Chr(34) & "0; url=" & basevirtualpath & _
exportfilename & Chr(34) & ">"
'clean up files that are more than 1 day old
 
set objFS = CreateObject("Scripting.FileSystemObject")
set objFC = objFS.GetFolder( Left( basePath, ( len( basePath ) - 1) ) )
set objF = objFC.Files
for each Item in objF
 testfilename = UCase(Item.Name)
 testextension = UCase(right(testfilename,4))
 if UCase(left( testfilename, len( "Orders report"))) = UCase("Orders report") and _
	testextension=UCase(fileextension) and _
	right(testfilename, 15) <> "-PARAMETERS.HTM" then
     if Item.DateCreated < Now - 1 then
         on error resume next
         Item.Delete
     end if
 end if  
next
set  objF = Nothing
set objFC = Nothing
set objFS = Nothing
%>
