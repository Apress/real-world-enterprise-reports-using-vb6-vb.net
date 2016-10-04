Dim objApplication As New CRAXDRT.Application
Dim objReport As New CRAXDRT.Report
Dim objParameterFields As CRAXDRT.ParameterFieldDefinitions
Dim objParameterField1 As CRAXDRT.ParameterFieldDefinition

' Dynamically add the CRVIEWER control to the form
Set CRViewer91 = Me.CRViewer91

'Set the report object
Set objReport = objApplication.OpenReport(App.Path & "\requestor.rpt")

objReport.DiscardSavedData

Set objParameterFields = objReport.ParameterFields

Set objParameterField1 = objParameterFields.Item(1)
objParameterField1.AddCurrentValue ("SC")
