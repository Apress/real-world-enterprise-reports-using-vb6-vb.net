VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} Labels 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14055
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24791
   _ExtentY        =   14499
   SectionData     =   "Labels.dsx":0000
End
Attribute VB_Name = "Labels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objDataControl As DDActiveReports2.DataControl
Dim objNameField As DDActiveReports2.Field
Dim objContactField As DDActiveReports2.Field
Dim objAddressField As DDActiveReports2.Field
Dim objCSZField As DDActiveReports2.Field

Private Sub ActiveReport_DataInitialize()
    'These are custom fields whose value will be in FetchData
    Me.Fields.Add "Contact"
    Me.Fields.Add "CSZ"
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    If EOF Then
        Exit Sub
    End If

    With objDataControl
        Me.Fields("Contact").Value = .Recordset("FirstName") & " " & .Recordset("LastName")
        Me.Fields("CSZ").Value = .Recordset("City") & ", " & .Recordset("State") & " " & .Recordset("Zip")
    End With

End Sub

Private Sub ActiveReport_ReportStart()

    Set objDataControl = Me.Sections("Detail").Controls.Add("DDActiveReports2.DataControl")
    objDataControl.Name = "adoRequester"

    objDataControl.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                   "Data Source=C:\BookCode\SampleDatabase.mdb;" & _
                                   "Persist Security Info=False"

    objDataControl.Source = "SELECT * FROM Requester ORDER BY Lastname"

    If IsSection(Me, "PageHeader") Then
        Call Me.Sections.Remove("PageHeader")
    End If

    If IsSection(Me, "PageFooter") Then
        Call Me.Sections.Remove("PageFooter")
    End If

    With Me.Sections("Detail")
        .ColumnCount = 2
        .ColumnDirection = ddCDDownAcross
        .KeepTogether = True
        .Height = 1440
    End With


    Set objContactField = Me.Sections("Detail").Controls.Add("DDActiveReports2.Field")

    With objContactField
        .Name = "fldContactName"
        .Height = 300
        .Width = 4500
        .Top = 300
        .Left = 0
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .DataField = "Contact"  'Custom Unbound Field
        .Border.Shadow = False
    End With


    Set objAddressField = Me.Sections("Detail").Controls.Add("DDActiveReports2.Field")

    With objAddressField
        .Name = "fldAddress1"
        .Height = 300
        .Width = 4500
        .Top = 600
        .Left = 0
        .CanShrink = True
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .DataField = "Address1"
        .Border.Shadow = False
    End With


    Set objAddressField = Me.Sections("Detail").Controls.Add("DDActiveReports2.Field")

    With objAddressField
        .Name = "fldAddress2"
        .Height = 300
        .Width = 4500
        .Top = 900
        .Left = 0
        .CanShrink = True
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .DataField = "Address2"
        .Border.Shadow = False
    End With


    Set objCSZField = Me.Sections("Detail").Controls.Add("DDActiveReports2.Field")

    With objCSZField
        .Name = "fldCity"
        .Height = 300
        .Width = 4500
        .Top = 1200
        .Left = 0
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
        .DataField = "CSZ"  'Custom Unbound Field
    End With

End Sub

Function IsSection(objReport As Object, cSection As String) As Boolean
    Dim bResult As Boolean
    Dim iCnt As Integer
    Dim x As Integer

    iCnt = objReport.Sections.Count - 1
    bResult = False

    For x = 0 To iCnt

        If objReport.Sections(x).Name = cSection Then
            bResult = True
            Exit For
        End If

    Next x

    IsSection = bResult

End Function



