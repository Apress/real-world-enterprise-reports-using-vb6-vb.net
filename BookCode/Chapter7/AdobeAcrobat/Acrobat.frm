VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   ClientHeight    =   11085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10320
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpenPDF 
      Cancel          =   -1  'True
      Caption         =   "Open PDF"
      Height          =   375
      Left            =   9840
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9840
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
      Height          =   375
      Left            =   9840
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   9840
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Last"
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First"
      Height          =   375
      Left            =   9840
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   9840
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objAcroExchPDDoc As CAcroPDDoc
Dim objAcroExchAVDoc As CAcroAVDoc
Dim iPage As Integer
Dim cPDFFileName As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFirst_Click()
        
    iPage = 0
    
    Me.Refresh
    
End Sub

Private Sub cmdLast_Click()
    iPage = objAcroExchPDDoc.GetNumPages - 1
    
    Me.Refresh
End Sub

Private Sub cmdNext_Click()
    
    If iPage < objAcroExchPDDoc.GetNumPages - 1 Then
        iPage = iPage + 1
    End If
    
    Me.Refresh
    
End Sub

Private Sub cmdOpenPDF_Click()
    Dim x As Integer
    
    iPage = 1
    
    With CommonDialog1
        .FileName = ""
        .DefaultExt = "pdf"
        .Filter = "PDF Files (*.pdf)|*.pdf"
        .ShowOpen
        cPDFFileName = .FileName
    End With
    
    Me.Caption = cPDFFileName
        
    Set objAcroExchPDDoc = CreateObject("AcroExch.PDDoc")
    x = objAcroExchPDDoc.Open(cPDFFileName)
    
    Me.Refresh
    
End Sub

Private Sub cmdPrevious_Click()

    If iPage > 0 Then
        iPage = iPage - 1
    End If
    
    Me.Refresh
    
End Sub

Private Sub cmdInfo_Click()
    Dim cMsg As String
    
    With objAcroExchPDDoc
        cMsg = "Title: " & .GetInfo("Title") & vbCrLf
        cMsg = cMsg & "Subject: " & .GetInfo("Subject") & vbCrLf
        cMsg = cMsg & "Author: " & .GetInfo("Author") & vbCrLf
        cMsg = cMsg & "Keywords: " & .GetInfo("Keywords") & vbCrLf
        cMsg = cMsg & "Creator: " & .GetInfo("Creator") & vbCrLf
        cMsg = cMsg & "Created: " & .GetInfo("Created") & vbCrLf
        cMsg = cMsg & "Modified: " & .GetInfo("Modified") & vbCrLf
        cMsg = cMsg & "Producer: " & .GetInfo("Producer")
    End With
    
    MsgBox cMsg

End Sub

Private Sub Form_Paint()
    Dim objAcroExchPDPage As CAcroPDPage
    Dim objAcroRect As CAcroRect
    Dim x As Integer
    
    If cPDFFileName = vbNullString Then
        Exit Sub
    End If
    
    Cls
    
    Set objAcroExchPDPage = objAcroExchPDDoc.AcquirePage(iPage)
    Set objAcroRect = CreateObject("AcroExch.Rect")

    With objAcroRect
        .Top = 792
        .bottom = 0
        .Left = 0
        .Right = 612
    End With
       
    x = objAcroExchPDPage.DrawEx(hWnd, 0, objAcroRect, 0, 0, 100)
    
    Set objAcroExchPDPage = Nothing
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set objAcroExchPDDoc = Nothing
    Set objAcroExchAVDoc = Nothing

End Sub
