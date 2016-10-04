VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Word-to-PDF Conversion"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpenDoc 
      Caption         =   "Open Document"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents objAdobe As PdfDistiller
Attribute objAdobe.VB_VarHelpID = -1

Private Sub cmdOpenDoc_Click()
    Dim cWordFileName As String

    With CommonDialog1
        .FileName = ""
        .DefaultExt = "pdf"
        .Filter = "Word DOC Files (*.doc)|*.doc"
        .ShowOpen
        cWordFileName = .FileName
    End With
    
    Me.Caption = cWordFileName
    
    Call ConvertDoc(cWordFileName)
            
End Sub

Private Sub ConvertDoc(cWordFileName As String)
    Dim objWord As Word.Application
    Dim objDocument As Word.Document
    Dim cCurrentPrinter As String
    Dim cTempPSDocName As String
    Dim cPDFDocName As String
    
    Screen.MousePointer = vbHourglass
    
    Set objAdobe = CreateObject("PdfDistiller.PdfDistiller.1")
    
    Set objWord = CreateObject("Word.Application")
    
    Set objDocument = objWord.Documents.Open(cWordFileName)
    
    cTempPSDocName = Replace(cWordFileName, ".doc", ".ps")
    cPDFDocName = Replace(cWordFileName, ".doc", ".pdf")
    
    cCurrentPrinter = objWord.ActivePrinter
    
    objWord.ActivePrinter = "Generic PostScript Printer"
    
    Call objDocument.PrintOut(0, 0, 0, cTempPSDocName)
    
    objWord.ActivePrinter = cCurrentPrinter
    
    objDocument.Close False
    Set objDocument = Nothing
    
    If Dir(cPDFDocName) <> vbNullString Then
        Kill cPDFDocName
    End If
    
    Call objAdobe.FileToPDF(cTempPSDocName, cPDFDocName, 0)
    
    If Dir(cTempPSDocName) <> vbNullString Then
        Kill cTempPSDocName
    End If
            
    objWord.Quit
    Set objWord = Nothing
    
    Set objAdobe = Nothing
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub objAdobe_OnPercentDone(ByVal nPercentDone As Long)
    ProgressBar1.Value = nPercentDone
End Sub

Private Sub objAdobe_OnJobDone(ByVal strInputPostScript As String, _
    ByVal strOutputPDF As String)
    MsgBox "All done!", vbOKOnly, Me.Caption
End Sub

Private Sub Form_Load()
    ProgressBar1.Value = 0
End Sub
