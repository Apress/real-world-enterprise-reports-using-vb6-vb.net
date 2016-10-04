Imports System
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document

Public Class Hyperlink
    Inherits ActiveReport
    Public Sub New()
        MyBase.New()
        InitializeReport()
    End Sub
#Region "ActiveReports Designer generated code"
    Private WithEvents PageHeader As PageHeader = Nothing
    Private WithEvents Detail As Detail = Nothing
    Private WithEvents PageFooter As PageFooter = Nothing
	Public ds As DataDynamics.ActiveReports.DataSources.SqlDBDataSource = Nothing
	Private Label1 As Label = Nothing
	Private Label2 As Label = Nothing
	Private Label3 As Label = Nothing
	Private txtCompanyName As TextBox = Nothing
	Private txtContactName As TextBox = Nothing
	Private txtCustomerID As TextBox = Nothing
	Public Sub InitializeReport()
		Me.LoadLayout(Me.GetType, "WindowsApplication5.HyperLink.rpx")
		Me.ds = CType(Me.DataSource,DataDynamics.ActiveReports.DataSources.SqlDBDataSource)
		Me.PageHeader = CType(Me.Sections("PageHeader"),DataDynamics.ActiveReports.PageHeader)
		Me.Detail = CType(Me.Sections("Detail"),DataDynamics.ActiveReports.Detail)
		Me.PageFooter = CType(Me.Sections("PageFooter"),DataDynamics.ActiveReports.PageFooter)
		Me.Label1 = CType(Me.PageHeader.Controls(0),DataDynamics.ActiveReports.Label)
		Me.Label2 = CType(Me.PageHeader.Controls(1),DataDynamics.ActiveReports.Label)
		Me.Label3 = CType(Me.PageHeader.Controls(2),DataDynamics.ActiveReports.Label)
		Me.txtCompanyName = CType(Me.Detail.Controls(0),DataDynamics.ActiveReports.TextBox)
		Me.txtContactName = CType(Me.Detail.Controls(1),DataDynamics.ActiveReports.TextBox)
		Me.txtCustomerID = CType(Me.Detail.Controls(2),DataDynamics.ActiveReports.TextBox)
	End Sub

#End Region


    Private Sub Hyperlink_ReportStart(ByVal sender As Object, _
                ByVal e As System.EventArgs) _
                Handles MyBase.ReportStart

        Dim objFont As New Font("Arial", 10, FontStyle.Underline)

        txtCompanyName.ForeColor = Color.Blue
        txtCompanyName.Font = objFont

    End Sub

    Private Sub Detail_BeforePrint(ByVal sender As Object, _
        ByVal e As System.EventArgs) _
        Handles Detail.BeforePrint

        txtCompanyName.HyperLink = txtCustomerID.Text

    End Sub


    Private Sub PageHeader_Format(ByVal sender As Object, ByVal e As System.EventArgs) Handles PageHeader.Format

    End Sub
End Class
