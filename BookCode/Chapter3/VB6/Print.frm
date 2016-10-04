VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrint 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Print Report"
   ClientHeight    =   1560
   ClientLeft      =   3135
   ClientTop       =   3210
   ClientWidth     =   3075
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1560
   ScaleWidth      =   3075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrSetup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Printer &Setup"
      Height          =   372
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1212
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destination"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton rbPrint 
         Caption         =   "Screen"
         Height          =   235
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton rbPrint 
         Caption         =   "Printer"
         Height          =   235
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1212
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objReport As Report

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim x As Integer
    Dim cDocName As String
    
    On Error GoTo PrintOKError
        
    For x = 0 To rbPrint.Count - 1
        If rbPrint(x) Then
            objReport.PrintDestination = x
            Exit For
        End If
    Next x
    
    
    Set frmReportViewer.objReport = objReport
    
    If x = SCREEN_VIEW Then
        frmReportViewer.Show vbModal
    Else
        Load frmReportViewer
    End If
        
Exit Sub

PrintOKError:
    Resume Next
End Sub

Private Sub cmdPrSetup_Click()

    With CommonDialog1
        .DialogTitle = "Printer Setup"
        .ShowPrinter
    End With
    
End Sub



