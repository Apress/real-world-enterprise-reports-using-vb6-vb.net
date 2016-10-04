<%

Option Explicit

Dim objSessionMgr, objEnterpriseSession, objService 
Dim objSchedulingInfo, objInfoObjects
Dim lReportID, cSQL

lReportID = 250

Set objSessionMgr = CreateObject("CrystalEnterprise.SessionMgr")

Set objEnterpriseSession = _
  objSessionMgr.Logon("administrator","","nydwetl3","secEnterprise")

Set objService = objEnterpriseSession.Service("", "InfoStore")

cSQL = "SELECT * " & _
	   "FROM CI_INFOOBJECTS " & _
	   "WHERE SI_ID = " & lReportID & " AND " & _
	   "SI_PROGID = 'CrystalEnterprise.Report'"

Set objInfoObjects = objService.Query(cSQL)

Set objSchedulingInfo = objInfoObjects.Item(1).SchedulingInfo

With Response
	.Write "<P> Report Name: " & objInfoObjects.Item(1).Properties("SI_NAME")
	.Write "<P> Report Description: " & objInfoObjects.Item(1).Description
	.Write "<P> Schedule Begin: " & objSchedulingInfo.BeginDate
	.Write "<P> Schedule End: " & objSchedulingInfo.EndDate
	.Write "<P> Frequency: " & objSchedulingInfo.Type
End With

%>
