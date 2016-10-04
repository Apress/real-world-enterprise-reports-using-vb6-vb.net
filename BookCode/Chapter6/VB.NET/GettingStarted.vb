Imports System
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document

Public Class GettingStarted
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
	Private TextBox1 As TextBox = Nothing
	Private TextBox2 As TextBox = Nothing
	Public Sub InitializeReport()
		Me.LoadLayout(Me.GetType, "WindowsApplication5.GettingStarted.rpx")
		Me.ds = CType(Me.DataSource,DataDynamics.ActiveReports.DataSources.SqlDBDataSource)
		Me.PageHeader = CType(Me.Sections("PageHeader"),DataDynamics.ActiveReports.PageHeader)
		Me.Detail = CType(Me.Sections("Detail"),DataDynamics.ActiveReports.Detail)
		Me.PageFooter = CType(Me.Sections("PageFooter"),DataDynamics.ActiveReports.PageFooter)
		Me.Label1 = CType(Me.PageHeader.Controls(0),DataDynamics.ActiveReports.Label)
		Me.Label2 = CType(Me.PageHeader.Controls(1),DataDynamics.ActiveReports.Label)
		Me.TextBox1 = CType(Me.Detail.Controls(0),DataDynamics.ActiveReports.TextBox)
		Me.TextBox2 = CType(Me.Detail.Controls(1),DataDynamics.ActiveReports.TextBox)
	End Sub

#End Region

End Class
