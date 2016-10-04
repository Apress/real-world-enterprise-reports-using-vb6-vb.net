Dim WithEvents CRViewer91 As CRVIEWER9LibCtl.CRViewer9
Dim objApplication As New CRAXDRT.Application
Dim objReport As New CRAXDRT.Report

Private Sub Form_Load()
    
    ' Dynamically add the CRVIEWER control to the form
    Set CRViewer91 = Me.CRViewer91
    
    'Set the report object
    Set objReport = objApplication.OpenReport(App.Path & "\requestor.rpt")

    objReport.DiscardSavedData

    With CRViewer91
        .DisplayGroupTree = False
        .EnableGroupTree = False
        .EnableNavigationControls = True
        .EnableSearchControl = True
        .EnableExportButton = True
        .EnableRefreshButton = False
        .ReportSource = objReport
        .ViewReport
    End With
    
    'maximizes the window
    Me.WindowState = vbMaximized
        
End Sub
