VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} FromScratch 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14055
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24791
   _ExtentY        =   14499
   SectionData     =   "FromScratch.dsx":0000
End
Attribute VB_Name = "FromScratch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
    Call CreateReport
End Sub

Sub CreateReport()
    Dim objDataControl As DDActiveReports2.DataControl
    Dim objField As DDActiveReports2.Field
    Dim objLabel As DDActiveReports2.Label
    
    'Create a data control and set the connection propeties
    Set objDataControl = Me.Sections("Detail").Controls.Add("DDActiveReports2.DataControl")

    With objDataControl
        .Name = "adoFreight"
    
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                            "Data Source=C:\Program Files\Microsoft Visual Studio\VB98\Nwind.mdb;" & _
                            "Persist Security Info=False"
    
        .Source = "SELECT c.CompanyName, i.OrderDate, i.ShippedDate, i.Freight " & _
                  "FROM Customers c LEFT JOIN Invoices I ON c.CustomerID = i.CustomerID " & _
                  "ORDER BY c.CompanyName, i.Freight DESC"
    End With
    
    'Add sections to the report
    Call Me.Sections.Add("ReportHeader1", 0, ddSTReportHeader, 400)
    Call Me.Sections.Add("PageHeader1", 1, ddSTPageHeader, 750)
    Call Me.Sections.Add("GroupHeader1", 2, ddSTGroupHeader, 750)
    Call Me.Sections.Add("GroupFooter1", 4, ddSTGroupFooter, 400)
    Call Me.Sections.Add("PageFooter1", 5, ddSTPageFooter, 400)
    Call Me.Sections.Add("ReportFooter1", 6, ddSTReportFooter, 400)
    
    Me.Sections("GroupHeader1").DataField = "CompanyName"
    
    Me.Sections("Detail").Height = 350

    Set objField = Me.Sections("GroupHeader1").Controls.Add("DDActiveReports2.Field")

    With objField
        .Name = "fldCompanyName"
        .DataField = "CompanyName"
        .Height = 300
        .Width = 3500
        .Top = 100
        .Left = 0
        .BackStyle = 0
        .Font.Bold = True
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
    End With
   

    Set objLabel = Me.Sections("GroupHeader1").Controls.Add("DDActiveReports2.Label")

    With objLabel
        .Name = "lblLabel1"
        .Caption = "Order Date"
        .Font.Italic = True
        .Height = 500
        .Width = 1500
        .Top = 450
        .Left = 0
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
    End With


    Set objLabel = Me.Sections("GroupHeader1").Controls.Add("DDActiveReports2.Label")

    With objLabel
        .Name = "lblLabel2"
        .Caption = "Ship Date"
        .Font.Italic = True
        .Height = 500
        .Width = 1500
        .Top = 450
        .Left = 1800
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
    End With


    Set objLabel = Me.Sections("GroupHeader1").Controls.Add("DDActiveReports2.Label")

    With objLabel
        .Name = "lblLabel3"
        .Caption = "Amount"
        .Alignment = ddTXRight
        .Font.Italic = True
        .Height = 500
        .Width = 1500
        .Top = 450
        .Left = 2600
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
    End With


    Set objField = Me.Sections("Detail").Controls.Add("DDActiveReports2.Field")

    With objField
        .Name = "fldOrderDate"
        .DataField = "OrderDate"
        .Height = 300
        .Width = 1500
        .Top = 100
        .Left = 0
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
    End With
    
    
    Set objField = Me.Sections("Detail").Controls.Add("DDActiveReports2.Field")

    With objField
        .Name = "fldShippedDate"
        .DataField = "ShippedDate"
        .Height = 300
        .Width = 1500
        .Top = 100
        .Left = 1800
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
    End With
    
    
    Set objField = Me.Sections("Detail").Controls.Add("DDActiveReports2.Field")

    With objField
        .Name = "fldFreight"
        .DataField = "Freight"
        .Alignment = ddTXRight
        .OutputFormat = "#,##0.00;(#,##0.00)"
        .Height = 300
        .Width = 1500
        .Top = 100
        .Left = 2600
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
    End With
    
    Set objField = Me.Sections("GroupFooter1").Controls.Add("DDActiveReports2.Field")
    
    With objField
        .Name = "fldSubTotal"
        .SummaryType = ddSMSubTotal
        .SummaryFunc = ddSFDSum
        .SummaryRunning = ddSRGroup
        .SummaryGroup = "GroupHeader1"
        .DataField = "Freight"
        .Alignment = ddTXRight
        .OutputFormat = "#,##0.00;(#,##0.00)"
        .Height = 300
        .Width = 1500
        .Top = 100
        .Left = 2600
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
    End With
    
    
    
    Set objLabel = Me.Sections("ReportFooter1").Controls.Add("DDActiveReports2.Label")
    
    With objLabel
        .Name = "lblLabel7"
        .Caption = "Grand Total:"
        .Font.Bold = True
        .Alignment = ddTXRight
        .Height = 300
        .Width = 1500
        .Top = 100
        .Left = 1500
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
    End With
    
    Set objField = Me.Sections("ReportFooter1").Controls.Add("DDActiveReports2.Field")
    
    With objField
        .Name = "fldGrandTotal"
        .SummaryType = ddSMGrandTotal
        .SummaryFunc = ddSFDSum
        .SummaryRunning = ddSRAll
        .DataField = "Freight"
        .Alignment = ddTXRight
        .OutputFormat = "#,##0.00;(#,##0.00)"
        .Font.Bold = True
        .Height = 300
        .Width = 1500
        .Top = 100
        .Left = 2600
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
    End With
    
End Sub

