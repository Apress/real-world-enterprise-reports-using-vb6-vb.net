Dim objApplication As New CRAXDRT.Application
Dim objReport As New CRAXDRT.Report
Dim objParameterFields As CRAXDRT.ParameterFieldDefinitions
Dim objParameterField1 As CRAXDRT.ParameterFieldDefinition

Sub RunReport(cFileName As String)
    
    Set objReport = objApplication.OpenReport(cFileName)
    
    objReport.DiscardSavedData
                
    Set objParameterFields = objReport.ParameterFields
    Set objParameterField1 = objParameterFields.Item(1)
    
    Call CreateLabel(1, 100, 100, 1800, 285, objParameterField1.Prompt)
    Call CreateComboBox(1, 100, 1900, 900, 285)
    
    Call LoadComboBox(1)
        
    objReport.PaperOrientation = crLandscape
    objReport.PaperSize = crPaperLegal
    objReport.ReportTitle = "Ex-presidents list to be printed on " & objReport.PrinterName
    objReport.PaperSource = crPRBinUpper
            
End Sub
