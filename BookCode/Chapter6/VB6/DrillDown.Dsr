VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} DrillDown 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14055
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24791
   _ExtentY        =   14499
   SectionData     =   "DrillDown.dsx":0000
End
Attribute VB_Name = "DrillDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
    fldCompanyName.ForeColor = vbBlue
    fldCompanyName.Font.Underline = True
End Sub

Private Sub Detail_BeforePrint()
    fldCompanyName.Hyperlink = fldCustomerID.Text
End Sub

Private Sub ActiveReport_hyperLink(ByVal Button As Integer, link As String)
    Dim cSQL As String
    Dim objReport As DrillDownDetail
    
    Set objReport = New DrillDownDetail
    
    cSQL = "SELECT OrderDate, ShippedDate, Freight " & _
           "FROM orders " & _
           "WHERE customerid = " & Chr(39) & link & Chr(39)
        
    objReport.DataControl1.Source = cSQL
    
    objReport.Show
    
End Sub

