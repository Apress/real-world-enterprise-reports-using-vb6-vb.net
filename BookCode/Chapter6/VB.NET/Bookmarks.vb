Imports System
Imports System.Data.SqlClient
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document
Imports DataDynamics.ActiveReports.Export

Public Class BookmarksReport
    Inherits ActiveReport
    Public Sub New()
        MyBase.New()
        InitializeReport()
    End Sub
#Region "ActiveReports Designer generated code"
    Private WithEvents PageHeader As PageHeader = Nothing
    Private WithEvents GroupHeader1 As GroupHeader = Nothing
    Private WithEvents Detail As Detail = Nothing
    Private WithEvents GroupFooter1 As GroupFooter = Nothing
    Private WithEvents PageFooter As PageFooter = Nothing
	Public ds As DataDynamics.ActiveReports.DataSources.SqlDBDataSource = Nothing
	Private Label1 As Label = Nothing
	Private Label2 As Label = Nothing
	Private Label3 As Label = Nothing
	Private fldCompanyName As TextBox = Nothing
	Private fldOrderDate As TextBox = Nothing
	Private fldShippedDate As TextBox = Nothing
	Private fldFreight As TextBox = Nothing
	Public Sub InitializeReport()
		Me.LoadLayout(Me.GetType, "WindowsApplication5.Bookmarks.rpx")
		Me.ds = CType(Me.DataSource,DataDynamics.ActiveReports.DataSources.SqlDBDataSource)
		Me.PageHeader = CType(Me.Sections("PageHeader"),DataDynamics.ActiveReports.PageHeader)
		Me.GroupHeader1 = CType(Me.Sections("GroupHeader1"),DataDynamics.ActiveReports.GroupHeader)
		Me.Detail = CType(Me.Sections("Detail"),DataDynamics.ActiveReports.Detail)
		Me.GroupFooter1 = CType(Me.Sections("GroupFooter1"),DataDynamics.ActiveReports.GroupFooter)
		Me.PageFooter = CType(Me.Sections("PageFooter"),DataDynamics.ActiveReports.PageFooter)
		Me.Label1 = CType(Me.GroupHeader1.Controls(0),DataDynamics.ActiveReports.Label)
		Me.Label2 = CType(Me.GroupHeader1.Controls(1),DataDynamics.ActiveReports.Label)
		Me.Label3 = CType(Me.GroupHeader1.Controls(2),DataDynamics.ActiveReports.Label)
		Me.fldCompanyName = CType(Me.GroupHeader1.Controls(3),DataDynamics.ActiveReports.TextBox)
		Me.fldOrderDate = CType(Me.Detail.Controls(0),DataDynamics.ActiveReports.TextBox)
		Me.fldShippedDate = CType(Me.Detail.Controls(1),DataDynamics.ActiveReports.TextBox)
		Me.fldFreight = CType(Me.Detail.Controls(2),DataDynamics.ActiveReports.TextBox)
	End Sub

#End Region

    Dim iExportOption As Integer

    Property ExportOption()
        Get
            ExportOption = iExportOption
        End Get
        Set(ByVal Value)
            iExportOption = Value
        End Set
    End Property


    Private Sub Detail_Format(ByVal sender As Object, _
        ByVal e As System.EventArgs) Handles Detail.Format

        ' Me.Detail.AddBookmark(fldCompanyName.Text + "\" + fldOrderDate.Text)
        Me.Detail.AddBookmark(fldCompanyName.Text)

    End Sub

    Private Sub Report_End()
        Dim objExcelExport As New Xls.ExcelExport()
        Dim objPDFExport As New Pdf.PdfExport()
        Dim objHTMLExport As New Html.HtmlExport()
        Dim objRTFExport As New Rtf.RtfExport()
        Dim objTextExport As New DataDynamics.ActiveReports.Export.Text.TextExport()
        Dim objTIFFExport As New Tiff.TiffExport()
        Dim cPath As String

        cPath = Application.StartupPath

        Select Case iExportOption

            Case ExportType.ExportExcel
                With objExcelExport
                    .GenPageBreaks = True
                    .MultiSheet = True
                    .MinColumnWidth = 30
                    .Export(Me.Document, cPath & "\export.xls")
                End With

            Case ExportType.ExportHTML
                With objHTMLExport
                    .MultiPage = True
                    .Title = "HTML Export Document"
                    .Export(Me.Document, cPath & "\export.html")
                End With

            Case ExportType.ExportPDF
                objPDFExport.Export(Me.Document, cPath & "\export.pdf")

            Case ExportType.ExportRTF
                objRTFExport.Export(Me.Document, cPath & "\export.rtf")

            Case ExportType.ExportText
                objTextExport.Export(Me.Document, cPath & "\export.txt")

            Case ExportType.ExportTIFF
                objTIFFExport.Export(Me.Document, cPath & "\export.tif")

        End Select

    End Sub

End Class
