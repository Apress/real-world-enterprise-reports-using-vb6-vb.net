VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmReportView 
   Caption         =   "Form1"
   ClientHeight    =   7824
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   7824
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7572
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8892
      lastProp        =   500
      _cx             =   15684
      _cy             =   13356
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmReportView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objReport As New CRAXDRT.Report

Dim WithEvents CRViewer91 As CRVIEWER9LibCtl.CRViewer9
Attribute CRViewer91.VB_VarHelpID = -1

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

Private Sub Form_Resize()
    
    With CRViewer91
        .Height = Me.Height - 100
        .Width = Me.Width - 100
    End With
        
End Sub

