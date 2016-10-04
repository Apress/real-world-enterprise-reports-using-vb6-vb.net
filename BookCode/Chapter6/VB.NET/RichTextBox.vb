Imports System
Imports System.Data.SqlClient
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document

Public Class RichTextBoxRpt
    Inherits ActiveReport
    Public Sub New()
        MyBase.New()
        InitializeReport()
    End Sub
#Region "ActiveReports Designer generated code"
    Private WithEvents Detail As Detail = Nothing
	Private RichTextBox1 As RichTextBox = Nothing
	Public Sub InitializeReport()
		Me.LoadLayout(Me.GetType, "WindowsApplication5.RichTextBox.rpx")
		Me.Detail = CType(Me.Sections("Detail"),DataDynamics.ActiveReports.Detail)
		Me.RichTextBox1 = CType(Me.Detail.Controls(0),DataDynamics.ActiveReports.RichTextBox)
	End Sub

#End Region

    Dim objDR As SqlDataReader
    Dim cCompanyName As String
    Dim cAddress As String
    Dim cCSZ As String

    Private Sub RichText_ReportStart(ByVal sender As Object, _
        ByVal e As System.EventArgs) Handles MyBase.ReportStart

        Dim objCommand As SqlCommand
        Dim objSqlConnection As New SqlConnection()
        Dim cSQL As String
        Dim objTextBox As TextBox

        'First, establish a data connection
        objSqlConnection.ConnectionString = "Data Source=(local);" & _
                                            "Initial catalog=Northwind;" & _
                                            "Integrated security=SSPI;" & _
                                            "Persist security info=False"
        objSqlConnection.Open()

        cSQL = "SELECT * " & _
               "FROM Customers "

        objCommand = New SqlCommand(cSQL, objSqlConnection)

        objDR = objCommand.ExecuteReader()
        objDR.Read()


        With RichTextBox1
            .RTF() = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033" & _
                     "{\fonttbl{\f0\fswiss\fcharset0 Arial;}" & _
                     "{\f1\fswiss\fprq2\fcharset0 Univers;}}" & _
                     "\viewkind4\uc1\pard\f0\fs20\par" & _
                     "\b [!CompanyName]\par" & _
                     "[!Address]\par" & _
                     "[!CSZ]\b0\par" & _
                     "\par" & _
                     "\f1 This is the reminder letter for [!CompanyName]\f0\par}"

            .CanGrow = True
            .Multiline = True

        End With

    End Sub

    Private Sub RichText_FetchData(ByVal sender As Object, _
        ByVal eArgs As DataDynamics.ActiveReports.ActiveReport.FetchEventArgs) _
        Handles MyBase.FetchData

        If eArgs.EOF Then
            Exit Sub
        End If

        cCompanyName = objDR.Item("CompanyName")
        cAddress = objDR.Item("Address")
        cCSZ = objDR.Item("City") & ", " & objDR.Item("Region") & " " & _
               objDR.Item("PostalCode")

    End Sub

    Private Sub RichText_Format(ByVal sender As Object, _
        ByVal e As System.EventArgs) Handles Detail.Format

        With RichTextBox1
            .ReplaceField("CompanyName", cCompanyName)
            .ReplaceField("Address", cAddress)
            .ReplaceField("CSZ", cCSZ)
        End With

    End Sub
End Class
