Public objReport As New CRAXDRT.Report

Dim WithEvents CRViewer91 As CRVIEWERLibCtl.CRViewer

Private Sub Form_Load()

    ' Dynamically add the CRVIEWER control to the form
    Set CRViewer91 = Me.CRViewer91
    
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

