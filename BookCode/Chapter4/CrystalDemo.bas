Attribute VB_Name = "Module1"
Option Explicit

Sub PrintReport()
    Dim objReport As New CRPEAuto.Report
    Set objReport = objCrystal.OpenReport(App.Path & "\Requestor.rpt")
     
     
'    Dim objCrystal As New CRPEAuto.Application
'    Dim CRXReportField As FieldObject
'
'
'    Dim objParams As CRPEAuto.ParameterFieldDefinition
'
'
    
    'Set CRXReportField = objReport.f
    
    'objReport.RecordSelectionFormula = "{product.ID} = 1"
    
    'objReport.PrintOut
'
'    Set objCrystal = Nothing
'    Set objReport = Nothing
    
End Sub
