Attribute VB_Name = "Module1"
Option Explicit

Sub StandardHeader(objReport As ActiveReport, cReportTitle As String)
    Dim objLabel As DDActiveReports2.Label
    Dim objPageCount As DDActiveReports2.SummaryType

    objReport.Sections.Add "PageHeader", 1, ddSTPageHeader, 500

    Set objLabel = objReport.Sections("PageHeader").Controls.Add("DDActiveReports2.Label")
    
    With objLabel
        .Name = "CompanyName"
        .Caption = "Seton Software Development, Inc."
        .Font.Bold = True
        .Font.Size = 9
        .Height = 300
        .Width = 4000
        .Top = 100
        .Left = 0
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
    End With

    Set objLabel = objReport.Sections("PageHeader").Controls.Add("DDActiveReports2.Label")

    With objLabel
        .Name = "Title"
        .Caption = cReportTitle
        .Font.Bold = True
        .Font.Size = 12
        .Height = 300
        .Width = 4000
        .Alignment = ddTXCenter
        .Top = 300
        .Left = 2000
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
    End With
    
End Sub


