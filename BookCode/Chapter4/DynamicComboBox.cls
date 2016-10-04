VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynamicComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents objComboBox As VB.ComboBox
Attribute objComboBox.VB_VarHelpID = -1

Public frmOwner As Form
Public iIndex As Integer

Public Sub CreateControl(iIndex As Integer, iTop As Integer, iLeft As Integer, iWidth As Integer, iHeight As Integer)

    Set objComboBox = frmOwner.Controls.Add("VB.ComboBox", "objComboBox" & iIndex, frmOwner)
    
    With objComboBox
        .Top = iTop
        .Left = iLeft
        .Width = iWidth
        .Visible = True
    End With
    
End Sub

Public Sub AddItem(vntDescr As Variant)

    objComboBox.AddItem vntDescr
    
End Sub

Public Function Text() As String
    Dim cResult As String
    
    cResult = objComboBox.Text
    
    Text = cResult
    
End Function



