Imports System
Imports System.Data.SqlClient
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document
Imports DataDynamics.ActiveReports.Export

Public Class UnboundReport
    Inherits ActiveReport
    Public Sub New()
        MyBase.New()
        InitializeReport()
    End Sub
#Region "ActiveReports Designer generated code"
    Private WithEvents ReportHeader As ReportHeader = Nothing
    Private WithEvents PageHeader As PageHeader = Nothing
    Private WithEvents Detail As Detail = Nothing
    Private WithEvents PageFooter As PageFooter = Nothing
    Private WithEvents ReportFooter As ReportFooter = Nothing
	Private Label3 As Label = Nothing
	Private Label1 As Label = Nothing
	Private Label2 As Label = Nothing
	Private fldCompanyName As TextBox = Nothing
	Private fldContactName As TextBox = Nothing
	Public Sub InitializeReport()
		Me.LoadLayout(Me.GetType, "WindowsApplication5.UnboundReport.rpx")
		Me.ReportHeader = CType(Me.Sections("ReportHeader"),DataDynamics.ActiveReports.ReportHeader)
		Me.PageHeader = CType(Me.Sections("PageHeader"),DataDynamics.ActiveReports.PageHeader)
		Me.Detail = CType(Me.Sections("Detail"),DataDynamics.ActiveReports.Detail)
		Me.PageFooter = CType(Me.Sections("PageFooter"),DataDynamics.ActiveReports.PageFooter)
		Me.ReportFooter = CType(Me.Sections("ReportFooter"),DataDynamics.ActiveReports.ReportFooter)
		Me.Label3 = CType(Me.ReportHeader.Controls(0),DataDynamics.ActiveReports.Label)
		Me.Label1 = CType(Me.PageHeader.Controls(0),DataDynamics.ActiveReports.Label)
		Me.Label2 = CType(Me.PageHeader.Controls(1),DataDynamics.ActiveReports.Label)
		Me.fldCompanyName = CType(Me.Detail.Controls(0),DataDynamics.ActiveReports.TextBox)
		Me.fldContactName = CType(Me.Detail.Controls(1),DataDynamics.ActiveReports.TextBox)
	End Sub

#End Region

    Dim objDR As SqlDataReader
    Dim iExportOption As Integer

    Property ExportOption()
        Get
            ExportOption = iExportOption
        End Get
        Set(ByVal Value)
            iExportOption = Value
        End Set
    End Property

    Private Sub UnboundReport_ReportStart(ByVal sender As Object, ByVal e As System.EventArgs) _
 Handles MyBase.ReportStart

        Dim objSqlConnection As New SqlConnection()
        Dim objCommand As SqlCommand
        Dim cSQL As String

        objSqlConnection.ConnectionString = "Data Source=(Local);" & _
                                            "Initial catalog=Northwind;" & _
                                            "Integrated security=SSPI;" & _
                                            "Persist security info=False"
        objSqlConnection.Open()

        cSQL = "SELECT CompanyName, ContactName " & _
               "FROM Customers "

        objCommand = New SqlCommand(cSQL, objSqlConnection)

        objDR = objCommand.ExecuteReader()
    End Sub

    Private Sub UnboundReport_ReportEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.ReportEnd

        Dim objExcelExport As New Xls.ExcelExport()
        Dim objPDFExport As New Pdf.PdfExport()
        Dim objHTMLExport As New Html.HtmlExport()
        Dim objRTFExport As New Rtf.RtfExport()
        Dim objTextExport As New DataDynamics.ActiveReports.Export.Text.TextExport()
        Dim objTIFFExport As New Tiff.TiffExport()
        Dim cPath As String
        Dim cFileName As String

        If iExportOption = -1 Then
            Exit Sub
        End If

        cPath = Application.StartupPath

        Select Case iExportOption

            Case ExportType.ExportExcel
                cFileName = cPath & "\export.xls"

                With objExcelExport
                    .GenPageBreaks = True
                    .MultiSheet = True
                    .MinColumnWidth = 30
                    .Export(Me.Document, cFileName)
                End With

            Case ExportType.ExportHTML
                cFileName = cPath & "\export.html"

                With objHTMLExport
                    .MultiPage = True
                    .Title = "HTML Export Document"
                    .Export(Me.Document, cFileName)
                End With

            Case ExportType.ExportPDF
                cFileName = cPath & "\export.pdf"

                objPDFExport.Export(Me.Document, cFileName)

            Case ExportType.ExportRTF
                cFileName = cPath & "\export.rtf"

                objRTFExport.Export(Me.Document, cFileName)

            Case ExportType.ExportText
                cFileName = cPath & "\export.txt"

                objTextExport.Export(Me.Document, cFileName)

            Case ExportType.ExportTIFF
                cFileName = cPath & "\export.tif"

                objTIFFExport.Export(Me.Document, cFileName)

        End Select

        MessageBox.Show("Sent to " + cFileName)

    End Sub

    Private Sub UnboundReport_DataInitialize(ByVal sender As Object, ByVal e As _
     System.EventArgs) Handles MyBase.DataInitialize
        Fields.Add("fldCompanyName")
        Fields.Add("fldContactName")
    End Sub

    Private Sub UnboundReport_FetchData(ByVal sender As Object, ByVal eArgs As DataDynamics.ActiveReports.ActiveReport.FetchEventArgs) Handles MyBase.FetchData
        Try
            objDR.Read()
            Me.Fields("fldCompanyName").Value = objDR("CompanyName").ToString()
            Me.Fields("fldContactName").Value = objDR("ContactName").ToString()

            eArgs.EOF = False
        Catch ex As Exception
            eArgs.EOF = True
        End Try
    End Sub

    Private Sub PageFooter_Format(ByVal sender As Object, ByVal e As System.EventArgs) Handles PageFooter.Format

    End Sub

    Private Sub PageHeader_Format(ByVal sender As Object, ByVal e As System.EventArgs) Handles PageHeader.Format

    End Sub

    Private Sub ReportHeader_Format(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReportHeader.Format

    End Sub
End Class
