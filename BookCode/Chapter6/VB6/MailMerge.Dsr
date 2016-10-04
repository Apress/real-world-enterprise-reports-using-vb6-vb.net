VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} MailMerge 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14055
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24791
   _ExtentY        =   14499
   SectionData     =   "MailMerge.dsx":0000
End
Attribute VB_Name = "MailMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
    RichEdit1.LoadFile App.Path & "\mailMergedoc.rtf", rtfRTF
    
    'LastName.CanShrink = True
End Sub

Private Sub Detail_Format()

    With RichEdit1
        .ReplaceField "LastName", "Ganz"
        .ReplaceField "FirstName", "Carl"
        .ReplaceField "Address1", "3 Barbieri Court"
        .ReplaceField "Address2", ""
        .ReplaceField "City", "Raritan"
        .ReplaceField "State", "NJ"
        .ReplaceField "Zip", "08869"
    End With
    
End Sub

