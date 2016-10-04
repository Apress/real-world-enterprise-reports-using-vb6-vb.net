Imports System
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document

Public Class MySubReport
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
	Private SubReport1 As SubReport = Nothing
	Public Sub InitializeReport()
		Me.LoadLayout(Me.GetType, "WindowsApplication5.SubReport.rpx")
		Me.ds = CType(Me.DataSource,DataDynamics.ActiveReports.DataSources.SqlDBDataSource)
		Me.PageHeader = CType(Me.Sections("PageHeader"),DataDynamics.ActiveReports.PageHeader)
		Me.Detail = CType(Me.Sections("Detail"),DataDynamics.ActiveReports.Detail)
		Me.PageFooter = CType(Me.Sections("PageFooter"),DataDynamics.ActiveReports.PageFooter)
		Me.SubReport1 = CType(Me.Detail.Controls(0),DataDynamics.ActiveReports.SubReport)
	End Sub

#End Region

    Dim cCustomerID As String

    Private Sub rptMain_FetchData(ByVal sender As Object, _
        ByVal e As DataDynamics.ActiveReports.ActiveReport.FetchEventArgs) _
        Handles MyBase.FetchData

        If e.EOF Then
            Exit Sub
        End If

        cCustomerID = Me.Fields("CustomerID").Value

    End Sub

    Private Sub Detail_Format(ByVal sender As Object, ByVal e As System.EventArgs) Handles Detail.Format
        Dim objReport As New ProductsOrderedSubRpt()
        Dim objDS As New DataSources.OleDBDataSource()
        Dim cSQL As String

        cSQL = "SELECT p.ProductName " & _
               "FROM Products p, Orders o, [Order Details] d " & _
               "WHERE(o.OrderID = d.OrderID) " & _
               "AND d.ProductID = p.ProductID " & _
               "AND o.CustomerID = " & Chr(39) & cCustomerID & Chr(39)

        objDS.ConnectionString = Me.ds.ConnectionString
        objDS.SQL = cSQL

        objReport.DataSource = objDS

        Me.SubReport1.Report = objReport

    End Sub
End Class
