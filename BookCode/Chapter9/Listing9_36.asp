<%

Option Explicit

Dim objSessionMgr, objEnterpriseSession, objService 
Dim objInfoObjects, objReportObject, objLogonTokenMgr 
Dim lReportID, cToken, cSQL, cRedirect

lReportID = 250

Set objSessionMgr = CreateObject("CrystalEnterprise.SessionMgr")

Set objEnterpriseSession = _
objSessionMgr.Logon("administrator","","nydwetl3","secEnterprise")

Set objService = objEnterpriseSession.Service("", "InfoStore")

Set objLogonTokenMgr = objEnterpriseSession.LogonTokenMgr
cToken = objLogonTokenMgr.CreateLogonToken("", 1, 100)

cSQL = "SELECT * " & _
   "FROM CI_INFOOBJECTS " & _
   "WHERE SI_ID = " & lReportID & " AND " & _
   "SI_PROGID = 'CrystalEnterprise.Report'"

Set objInfoObjects = objService.Query(cSQL)

Set objReportObject = objInfoObjects.Item(1).PluginInterface

With objReportObject.ReportParameters.Item(1).CurrentValues.Item(1)	
    .Value = 6164
End With

With objReportObject.ReportLogons.Item(1)
    .ServerName="MYSERVER"
    .Databasename="testdb"
    .UserName="myUID"
    .Password="mypass"
End With

objService.Commit objInfoObjects

cRedirect = "http://nydwetl3/crystal/enterprise/" & _
"admin/en/viewrpt.cwr?id=" & lReportID & _
"&apstoken=" & cToken

Response.Redirect(cRedirect)

set objSessionMgr = nothing
set objEnterpriseSession = nothing
set objService = nothing
set objInfoObjects = nothing
set objReportObject = nothing
set objLogonTokenMgr = nothing

%>
