VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} SubReport 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14040
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24765
   _ExtentY        =   14499
   SectionData     =   "SubReport.dsx":0000
End
Attribute VB_Name = "SubReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
    Dim objReport As FirstReport
    
    lblIntro.Caption = "This demo illustrates the cool things you can do with " & _
                       "subreports. If this were a real report you would see some " & _
                       "meangingful text here."
        
    lblExplain.Caption = "Down here is the same subreport with a diffrent " & _
                         "parameter passed to it "
        
    Set objReport = New FirstReport
    
    objReport.DataControl1.Source = "SELECT * FROM Requester WHERE State = 'CA'"
    Set SubReport1.object = objReport
    
    Set objReport = New FirstReport
        
    objReport.DataControl1.Source = "SELECT * FROM Requester WHERE State = 'NJ'"
    Set SubReport2.object = objReport

End Sub

Private Sub PageHeader_Format()

End Sub
