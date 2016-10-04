VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} Grouping 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   14052
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24786
   _ExtentY        =   14499
   SectionData     =   "Grouping.dsx":0000
End
Attribute VB_Name = "Grouping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
    Call StandardHeader(Me, "Grouping example")
    'Call CreateReport
End Sub

Private Sub Detail_Format()
    Me.TOC.Add GroupHeader1.GroupValue & "\" & fldOrderDate
End Sub

Private Sub GroupFooter1_Format()

    If fldCompanySubtotal.DataValue > 1000 Then
        fldCompanySubtotal.ForeColor = vbRed
    Else
        fldCompanySubtotal.ForeColor = vbBlack
    End If
    
End Sub

Private Sub GroupHeader1_Format()
    Me.TOC.Add GroupHeader1.GroupValue
End Sub

