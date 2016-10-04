VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ActiveReport1 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14055
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24791
   _ExtentY        =   14499
   SectionData     =   "ActiveReport1.dsx":0000
End
Attribute VB_Name = "ActiveReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()

    With Me.PageSettings
        .Orientation = ddOPortrait
        .PaperSize = 1
        .PaperBin = 1
        .BottomMargin = 1440
        .TopMargin = 1440
        .LeftMargin = 720
        .RightMargin = 720
    End With
    
End Sub

Private Sub ActiveReport_PageStart()
    If Me.Pages.Count > 0 Then
        Me.PageSettings.PaperBin = 2
    End If
End Sub


