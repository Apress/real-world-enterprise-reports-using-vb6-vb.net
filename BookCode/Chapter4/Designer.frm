VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{F8A2F63D-BF23-4B42-928C-CCB4545286D6}#9.2#0"; "CRDesignerCtrl.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDesigner 
   Caption         =   "Crystal Interactive Designer"
   ClientHeight    =   8232
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10416
   LinkTopic       =   "Form1"
   ScaleHeight     =   8232
   ScaleWidth      =   10416
   StartUpPosition =   3  'Windows Default
   Begin CRDESIGNERCTRLLibCtl.CRDesignerCtrl CRDesignerCtrl1 
      Height          =   7620
      Left            =   120
      OleObjectBlob   =   "Designer.frx":0000
      TabIndex        =   5
      Top             =   120
      Width           =   10200
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10215
      lastProp        =   500
      _cx             =   18018
      _cy             =   13361
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdDesignView 
      Caption         =   "View"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   7800
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpenExisting 
      Caption         =   "Open Existing Report"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdCreateNew 
      Caption         =   "Create New Report"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   7800
      Width           =   1935
   End
End
Attribute VB_Name = "frmDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objApplication As New CRAXDDRT.Application
Dim objReport As CRAXDDRT.Report
Dim cFileName As String

Dim WithEvents objSectionGF1a As CRAXDDRT.Section
Attribute objSectionGF1a.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCreateNew_Click()

    Set objReport = objApplication.NewReport

    With CRDesignerCtrl1
        .ReportObject = objReport
        .DisplayToolbar = True
        .DisplayGrid = True
    End With

End Sub

Private Sub cmdDesignView_Click()

    If cmdDesignView.Caption = "Design" Then
        cmdDesignView.Caption = "View"
        CRDesignerCtrl1.Visible = True
        CRViewer1.Visible = False
    Else
        cmdDesignView.Caption = "Design"
        CRDesignerCtrl1.Visible = False
        CRViewer1.Visible = True

        If Not objReport Is Nothing Then
            CRViewer1.ReportSource = objReport
            CRViewer1.ViewReport
        End If

    End If

End Sub

Private Sub cmdOpenExisting_Click()

    With CommonDialog1
       .Filter = "Crystal Reports|*.rpt"
       .ShowOpen
    End With

    cFileName = CommonDialog1.FileName

    If cFileName = vbNullString Then
       Exit Sub
    End If

    Set objReport = objApplication.OpenReport(cFileName, 1)

    CRDesignerCtrl1.ReportObject = objReport

End Sub

Private Sub Form_Resize()

    If Me.Height - cmdCreateNew.Height - 700 < 0 Then
        Exit Sub
    End If

    CRDesignerCtrl1.Width = Me.Width - 300
    CRDesignerCtrl1.Height = Me.Height - cmdCreateNew.Height - 700

    CRViewer91.Width = Me.Width - 300
    CRViewer91.Height = Me.Height - cmdCreateNew.Height - 700

    cmdCreateNew.Top = Me.Height - cmdCreateNew.Height - 500
    cmdOpenExisting.Top = Me.Height - cmdOpenExisting.Height - 500
    cmdDesignView.Top = Me.Height - cmdDesignView.Height - 500
    cmdCancel.Top = Me.Height - cmdCancel.Height - 500

End Sub

Private Sub objSectionGF1a_Format(ByVal pFormattingInfo As Object)
    Dim objFieldObject As CRAXDDRT.FieldObject

    Set objFieldObject = objReport.Sections("GF1a").ReportObjects.Item(1)

    If objFieldObject.Value > 200 Then
        objFieldObject.BackColor = vbRed
    Else
        objFieldObject.BackColor = vbWhite
    End If

End Sub
