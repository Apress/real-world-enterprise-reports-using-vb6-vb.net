VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form frmARViewer 
   Caption         =   "Report Viewer"
   ClientHeight    =   10785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13650
   LinkTopic       =   "Form1"
   ScaleHeight     =   10785
   ScaleWidth      =   13650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   10320
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   10320
      Width           =   975
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer21 
      Height          =   10095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   17806
      SectionData     =   "ARViewer.frx":0000
   End
End
Attribute VB_Name = "frmARViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    ARViewer21.PrintReport
End Sub

Private Sub Form_Load()
    ARViewer21.TOCVisible = True
    ARViewer21.TOCWidth = 3500
    ARViewer21.RulerVisible = False
    Set ARViewer21.ReportSource = Grouping
    
    
End Sub
