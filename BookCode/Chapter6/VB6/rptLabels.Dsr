VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptLabels 
   Caption         =   "Labels Report"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13635
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24051
   _ExtentY        =   16193
   SectionData     =   "rptLabels.dsx":0000
End
Attribute VB_Name = "rptLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objDataControl As DDActiveReports2.DataControl
Dim objNameField As DDActiveReports2.Field
Dim objContactField As DDActiveReports2.Field
Dim objAddressField As DDActiveReports2.Field
Dim objCityField As DDActiveReports2.Field
Dim objRegionField As DDActiveReports2.Field
Dim objCountryField As DDActiveReports2.Field

Private Sub ActiveReport_DataInitialize()
    Me.Fields.Add "Contact" ' This is a custom field will be set in FetchData
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    ' FetchData is typically used to set values of unbound fields not controls
    ' that are created in the DataInitialize event.
    If EOF Then
        Exit Sub
    End If

    With objDataControl
        ' the value will automatically be boung to the control as if it was one of the original
        ' recordset fields
        Me.Fields("Contact").Value = .Recordset("ContactName") & " [" & .Recordset("ContactTitle") & "]"
    End With

End Sub

Private Sub ActiveReport_ReportStart()
    Dim objField As DDActiveReports2.Field

    Set objDataControl = Me.Sections("Detail").Controls.Add("DDActiveReports2.DataControl")
    objDataControl.Name = "adoRequester"

    objDataControl.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                   "Data Source=C:\NWind.mdb;" & _
                                   "Persist Security Info=False"

    objDataControl.Source = "SELECT * FROM Customers"

    If IsSection(Me, "PageHeader") Then
        Call Me.Sections.Remove("PageHeader")
    End If

    If IsSection(Me, "PageFooter") Then
        Call Me.Sections.Remove("PageFooter")
    End If

    With Me.Sections("Detail")
        .ColumnCount = 2
        .ColumnDirection = ddCDAcrossDown
        .KeepTogether = True
        .Height = 1850
    End With


    Set objNameField = Me.Sections("Detail").Controls.Add("DDActiveReports2.Field")

    With objNameField
        .Name = "fldCompanyName"
        .Height = 300
        .Width = 4500
        .Top = 0
        .Left = 0
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
        .DataField = "CompanyName"
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
        .DataField = "Contact"  ' Custom Unbound Field
        .Border.Shadow = False
    End With


    Set objAddressField = Me.Sections("Detail").Controls.Add("DDActiveReports2.Field")

    With objAddressField
        .Name = "fldAddress"
        .Height = 300
        .Width = 4500
        .Top = 600
        .Left = 0
        .CanShrink = True
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .DataField = "Address"
        .Border.Shadow = False
    End With


    Set objCityField = Me.Sections("Detail").Controls.Add("DDActiveReports2.Field")

    With objCityField
        .Name = "fldCity"
        .Height = 300
        .Width = 4500
        .Top = 900
        .Left = 0
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
        .DataField = "City"
    End With

    Set objRegionField = Me.Sections("Detail").Controls.Add("DDActiveReports2.Field")

    With objRegionField
        .Name = "fldRegion"
        .Height = 300
        .Width = 4500
        .Top = 1200
        .Left = 0
        .BackStyle = 1
        .BackColor = &HC0FFFF
        .ForeColor = vbBlack
        .Border.Shadow = False
        .DataField = "Region"   ' Some values are blank, this should shrink
        .CanShrink = True
    End With

    Set objCountryField = Me.Sections("Detail").Controls.Add("DDActiveReports2.Field")

    With objCountryField
        .Name = "fldCountry"
        .Height = 300
        .Width = 4500
        .Top = 1500
        .Left = 0
        .BackStyle = 0
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .Border.Shadow = False
        .DataField = "Country"
    End With

End Sub

Function IsSection(objReport As ActiveReport, cSection As String) As Boolean
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


