VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} FirstReport 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   14052
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24786
   _ExtentY        =   14499
   SectionData     =   "FirstReport.dsx":0000
End
Attribute VB_Name = "FirstReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iLine As Integer

Private Sub ActiveReport_Initialize()
    DataControl1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BookCode\SampleDatabase.mdb;Persist Security Info=False"
    DataControl1.Source = "SELECT * FROM Requester"
End Sub

Private Sub ActiveReport_ReportStart()
    Dim iLabelCnt As Integer
    Dim x As Integer
           
    iLabelCnt = Me.Sections("PageHeader").Controls.Count - 1
            
        For x = 0 To iLabelCnt
        
        Me.Sections("PageHeader").Controls(x).Font.Bold = True
        
    Next x
    
End Sub

Private Sub Detail_Format()
    
    If (iLine Mod 2) = 0 Then
        Detail.BackStyle = ddBKNormal
        Detail.BackColor = &H8000000F
    Else
        Detail.BackStyle = ddBKTransparent
    End If
    
    iLine = iLine + 1
    
End Sub

