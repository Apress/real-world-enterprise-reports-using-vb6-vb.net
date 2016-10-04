VERSION 5.00
Begin VB.Form frmCriteria 
   Caption         =   "Criteria"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6216
   LinkTopic       =   "Form2"
   ScaleHeight     =   1680
   ScaleWidth      =   6216
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
         
Dim objApplication As New CRAXDRT.Application
Dim objReport As New CRAXDRT.Report
Dim objParameterFields As CRAXDRT.ParameterFieldDefinitions
Dim objParameterField1 As CRAXDRT.ParameterFieldDefinition
    
Dim colLabels As New Collection
Dim colComboBoxes As New Collection
     
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Call RunReport(App.Path & "\requestor.rpt")

End Sub

Sub RunReport(cFileName As String)
    
    Set objReport = objApplication.OpenReport(cFileName)
    
    objReport.DiscardSavedData
                
    Set objParameterFields = objReport.ParameterFields
    Set objParameterField1 = objParameterFields.Item(1)
    
    Call CreateLabel(1, 100, 100, 1800, 285, objParameterField1.Prompt)
    Call CreateComboBox(1, 100, 1900, 900, 285)
    
    Call LoadComboBox(1)
        
    objReport.PaperOrientation = crLandscape
    objReport.PaperSize = crPaperLegal
    objReport.ReportTitle = "Ex-presidents list to be printed on " & objReport.PrinterName
    objReport.PaperSource = crPRBinUpper
            
End Sub

Sub CreateComboBox(iIndex As Integer, iTop As Integer, iLeft As Integer, iWidth As Integer, iHeight As Integer)
    Dim objComboBox As DynamicComboBox
    
    Set objComboBox = New DynamicComboBox
    Set objComboBox.frmOwner = Me
    objComboBox.iIndex = iIndex
    Call objComboBox.CreateControl(iIndex, iTop, iLeft, iWidth, iHeight)
    colComboBoxes.Add objComboBox
End Sub

Sub CreateLabel(iIndex As Integer, iTop As Integer, iLeft As Integer, iWidth As Integer, iHeight As Integer, cCaption As String)
    Dim objLabel As DynamicLabel
    
    Set objLabel = New DynamicLabel
    Set objLabel.frmOwner = Me
    objLabel.iIndex = iIndex
    Call objLabel.CreateControl(iIndex, iTop, iLeft, iWidth, iHeight, cCaption)
    colLabels.Add objLabel
End Sub

Sub LoadComboBox(iIndex As Integer)
    
    With colComboBoxes(iIndex)
        Call .AddItem("CA")
        Call .AddItem("NJ")
        Call .AddItem("NY")
        Call .AddItem("SC")
    End With
    
End Sub

Private Sub cmdOK_Click()
    
    Set objParameterFields = objReport.ParameterFields
    
    Set objParameterField1 = objParameterFields.Item(1)
    Call objParameterField1.AddCurrentValue(colComboBoxes(1).Text)
    
    With frmReportView
        Set .objReport = objReport
        .Show vbModal
    End With
    
End Sub


