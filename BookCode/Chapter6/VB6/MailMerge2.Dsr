VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} MailMerge2 
   Caption         =   "MailMerge2"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14040
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24765
   _ExtentY        =   14499
   SectionData     =   "MailMerge2.dsx":0000
End
Attribute VB_Name = "MailMerge2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_FetchData(EOF As Boolean)

    If EOF Then
        Exit Sub
    End If

    With DataControl1
        fldName.DataValue = "" & .Recordset("FirstName") & " " & DataControl1.Recordset("LastName")
        fldAddress1.DataValue = "" & .Recordset("Address1")
        fldAddress2.DataValue = "" & .Recordset("Address2")
        fldCSZ.DataValue = "" & .Recordset("City") & ", " & .Recordset("State") & " " & .Recordset("Zip")
        
        fldSalutation.DataValue = "Dear " & .Recordset("FirstName")
    End With
    
End Sub

Private Sub ActiveReport_ReportStart()
    RichEdit1.LoadFile App.Path & "\mailmergedoc2.rtf", rtfRTF
    
    fldAddress2.CanShrink = True
End Sub

Private Sub Detail_Format()
    
    With RichEdit1
        .ReplaceField "ID", DataControl1.Recordset("ID")
    End With
    
    Detail.NewPage = ddNPAfter
    
End Sub

