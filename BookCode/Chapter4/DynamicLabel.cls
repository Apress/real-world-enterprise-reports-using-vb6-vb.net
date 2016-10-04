VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynamicLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents objLabel As VB.Label
Attribute objLabel.VB_VarHelpID = -1

Public frmOwner As Form
Public iIndex As Integer

Public Sub CreateControl(iIndex As Integer, iTop As Integer, iLeft As Integer, iWidth As Integer, iHeight As Integer, cCaption As String)
    
    Set objLabel = frmOwner.Controls.Add("VB.Label", "objLabel" & iIndex, frmOwner)
    
    With objLabel
        .Top = iTop
        .Left = iLeft
        .Width = iWidth
        .Height = iHeight
        .Caption = cCaption
        .Visible = True
    End With
    
End Sub

