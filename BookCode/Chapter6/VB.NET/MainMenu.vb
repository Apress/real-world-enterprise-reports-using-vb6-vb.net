Imports System.Drawing.Printing

Public Enum ExportType
    ExportExcel = 0
    ExportPDF = 1
    ExportHTML = 2
    ExportRTF = 3
    ExportText = 4
    ExportTIFF = 5
End Enum

Public Class frmMainMenu
    Inherits System.Windows.Forms.Form

    Dim iExport As Integer

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdBookmarks As System.Windows.Forms.Button
    Friend WithEvents cmdDynamic As System.Windows.Forms.Button
    Friend WithEvents cmdUnbound As System.Windows.Forms.Button
    Friend WithEvents cmdGettingStarted As System.Windows.Forms.Button
    Friend WithEvents cmdExport As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rbExportExcel As System.Windows.Forms.RadioButton
    Friend WithEvents rbExportPDF As System.Windows.Forms.RadioButton
    Friend WithEvents rbExportHTML As System.Windows.Forms.RadioButton
    Friend WithEvents rbExportRTF As System.Windows.Forms.RadioButton
    Friend WithEvents rbExportText As System.Windows.Forms.RadioButton
    Friend WithEvents rbExportTIFF As System.Windows.Forms.RadioButton
    Friend WithEvents cmdRichTextBox As System.Windows.Forms.Button
    Friend WithEvents cmdSubReport As System.Windows.Forms.Button
    Friend WithEvents cmdHyperLink As System.Windows.Forms.Button
    Friend WithEvents cmdInstalledPrinters As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdDynamic = New System.Windows.Forms.Button()
        Me.cmdUnbound = New System.Windows.Forms.Button()
        Me.cmdGettingStarted = New System.Windows.Forms.Button()
        Me.cmdBookmarks = New System.Windows.Forms.Button()
        Me.cmdExport = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbExportTIFF = New System.Windows.Forms.RadioButton()
        Me.rbExportText = New System.Windows.Forms.RadioButton()
        Me.rbExportRTF = New System.Windows.Forms.RadioButton()
        Me.rbExportHTML = New System.Windows.Forms.RadioButton()
        Me.rbExportPDF = New System.Windows.Forms.RadioButton()
        Me.rbExportExcel = New System.Windows.Forms.RadioButton()
        Me.cmdRichTextBox = New System.Windows.Forms.Button()
        Me.cmdSubReport = New System.Windows.Forms.Button()
        Me.cmdHyperLink = New System.Windows.Forms.Button()
        Me.cmdInstalledPrinters = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdDynamic
        '
        Me.cmdDynamic.Location = New System.Drawing.Point(10, 49)
        Me.cmdDynamic.Name = "cmdDynamic"
        Me.cmdDynamic.Size = New System.Drawing.Size(215, 30)
        Me.cmdDynamic.TabIndex = 0
        Me.cmdDynamic.Text = "Dynamic Report"
        '
        'cmdUnbound
        '
        Me.cmdUnbound.Location = New System.Drawing.Point(10, 89)
        Me.cmdUnbound.Name = "cmdUnbound"
        Me.cmdUnbound.Size = New System.Drawing.Size(215, 29)
        Me.cmdUnbound.TabIndex = 1
        Me.cmdUnbound.Text = "Unbound Report"
        '
        'cmdGettingStarted
        '
        Me.cmdGettingStarted.Location = New System.Drawing.Point(10, 10)
        Me.cmdGettingStarted.Name = "cmdGettingStarted"
        Me.cmdGettingStarted.Size = New System.Drawing.Size(215, 29)
        Me.cmdGettingStarted.TabIndex = 2
        Me.cmdGettingStarted.Text = "Getting Started"
        '
        'cmdBookmarks
        '
        Me.cmdBookmarks.Location = New System.Drawing.Point(10, 128)
        Me.cmdBookmarks.Name = "cmdBookmarks"
        Me.cmdBookmarks.Size = New System.Drawing.Size(215, 30)
        Me.cmdBookmarks.TabIndex = 3
        Me.cmdBookmarks.Text = "Bookmarks"
        '
        'cmdExport
        '
        Me.cmdExport.Location = New System.Drawing.Point(10, 168)
        Me.cmdExport.Name = "cmdExport"
        Me.cmdExport.Size = New System.Drawing.Size(215, 29)
        Me.cmdExport.TabIndex = 4
        Me.cmdExport.Text = "Export"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.rbExportTIFF, Me.rbExportText, Me.rbExportRTF, Me.rbExportHTML, Me.rbExportPDF, Me.rbExportExcel})
        Me.GroupBox1.Location = New System.Drawing.Point(240, 144)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(264, 64)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Export Format"
        '
        'rbExportTIFF
        '
        Me.rbExportTIFF.Location = New System.Drawing.Point(200, 40)
        Me.rbExportTIFF.Name = "rbExportTIFF"
        Me.rbExportTIFF.Size = New System.Drawing.Size(62, 20)
        Me.rbExportTIFF.TabIndex = 5
        Me.rbExportTIFF.Text = "TIFF"
        '
        'rbExportText
        '
        Me.rbExportText.Location = New System.Drawing.Point(200, 24)
        Me.rbExportText.Name = "rbExportText"
        Me.rbExportText.Size = New System.Drawing.Size(62, 20)
        Me.rbExportText.TabIndex = 4
        Me.rbExportText.Text = "Text"
        '
        'rbExportRTF
        '
        Me.rbExportRTF.Location = New System.Drawing.Point(104, 40)
        Me.rbExportRTF.Name = "rbExportRTF"
        Me.rbExportRTF.Size = New System.Drawing.Size(62, 19)
        Me.rbExportRTF.TabIndex = 3
        Me.rbExportRTF.Text = "RTF"
        '
        'rbExportHTML
        '
        Me.rbExportHTML.Location = New System.Drawing.Point(104, 24)
        Me.rbExportHTML.Name = "rbExportHTML"
        Me.rbExportHTML.Size = New System.Drawing.Size(82, 20)
        Me.rbExportHTML.TabIndex = 2
        Me.rbExportHTML.Text = "HTML"
        '
        'rbExportPDF
        '
        Me.rbExportPDF.Location = New System.Drawing.Point(10, 40)
        Me.rbExportPDF.Name = "rbExportPDF"
        Me.rbExportPDF.Size = New System.Drawing.Size(82, 20)
        Me.rbExportPDF.TabIndex = 1
        Me.rbExportPDF.Text = "PDF"
        '
        'rbExportExcel
        '
        Me.rbExportExcel.Checked = True
        Me.rbExportExcel.Location = New System.Drawing.Point(10, 24)
        Me.rbExportExcel.Name = "rbExportExcel"
        Me.rbExportExcel.Size = New System.Drawing.Size(82, 19)
        Me.rbExportExcel.TabIndex = 0
        Me.rbExportExcel.TabStop = True
        Me.rbExportExcel.Text = "Excel"
        '
        'cmdRichTextBox
        '
        Me.cmdRichTextBox.Location = New System.Drawing.Point(10, 207)
        Me.cmdRichTextBox.Name = "cmdRichTextBox"
        Me.cmdRichTextBox.Size = New System.Drawing.Size(215, 30)
        Me.cmdRichTextBox.TabIndex = 6
        Me.cmdRichTextBox.Text = "Rich Text Box"
        '
        'cmdSubReport
        '
        Me.cmdSubReport.Location = New System.Drawing.Point(10, 247)
        Me.cmdSubReport.Name = "cmdSubReport"
        Me.cmdSubReport.Size = New System.Drawing.Size(215, 29)
        Me.cmdSubReport.TabIndex = 7
        Me.cmdSubReport.Text = "SubReports"
        '
        'cmdHyperLink
        '
        Me.cmdHyperLink.Location = New System.Drawing.Point(8, 288)
        Me.cmdHyperLink.Name = "cmdHyperLink"
        Me.cmdHyperLink.Size = New System.Drawing.Size(215, 29)
        Me.cmdHyperLink.TabIndex = 8
        Me.cmdHyperLink.Text = "HyperLink"
        '
        'cmdInstalledPrinters
        '
        Me.cmdInstalledPrinters.Location = New System.Drawing.Point(8, 328)
        Me.cmdInstalledPrinters.Name = "cmdInstalledPrinters"
        Me.cmdInstalledPrinters.Size = New System.Drawing.Size(215, 29)
        Me.cmdInstalledPrinters.TabIndex = 9
        Me.cmdInstalledPrinters.Text = "Installed Printers"
        '
        'frmMainMenu
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(512, 360)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdInstalledPrinters, Me.cmdHyperLink, Me.cmdSubReport, Me.cmdRichTextBox, Me.GroupBox1, Me.cmdExport, Me.cmdBookmarks, Me.cmdGettingStarted, Me.cmdUnbound, Me.cmdDynamic})
        Me.Name = "frmMainMenu"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Active Reports for .NET"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdGroupReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles cmdBookmarks.Click, cmdUnbound.Click, cmdDynamic.Click, cmdGettingStarted.Click, _
                cmdExport.Click, cmdRichTextBox.Click, cmdSubReport.Click, cmdHyperLink.Click

        Dim frmForm As New frmReportViewer()
        Dim frmDynamicReport As New frmDynamicReport()

        If sender Is cmdGettingStarted Then
            frmForm.Report = New GettingStarted()
            frmForm.Show()

        ElseIf sender Is cmdDynamic Then
            frmDynamicReport.Show()

        ElseIf sender Is cmdUnbound Then
            Dim objReport As New UnboundReport()

            objReport.ExportOption = -1

            frmForm.Report = objReport
            frmForm.Show()

        ElseIf sender Is cmdBookmarks Then
            frmForm.Report = New BookmarksReport()
            frmForm.Show()

        ElseIf sender Is cmdExport Then
            Dim objReport As New UnboundReport()

            objReport.ExportOption = iExport

            frmForm.Report = objReport
            frmForm.Show()

        ElseIf sender Is cmdRichTextBox Then
            frmForm.Report = New RichTextBoxRpt()
            frmForm.Show()

        ElseIf sender Is cmdSubReport Then
            frmForm.Report = New MySubReport()
            frmForm.Show()

        ElseIf sender Is cmdHyperLink Then
            frmForm.Report = New Hyperlink()
            frmForm.Show()

        End If

    End Sub


    Private Sub rbExport_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles rbExportExcel.CheckedChanged, rbExportPDF.CheckedChanged, _
                rbExportHTML.CheckedChanged, rbExportRTF.CheckedChanged, _
                rbExportText.CheckedChanged, rbExportTIFF.CheckedChanged

        Dim objRadioButton As RadioButton

        objRadioButton = CType(sender, RadioButton)

        Select Case objRadioButton.Name

            Case "rbExportExcel"
                iExport = ExportType.ExportExcel

            Case "rbExportPDF"
                iExport = ExportType.ExportPDF

            Case "rbExportHTML"
                iExport = ExportType.ExportHTML

            Case "rbExportRTF"
                iExport = ExportType.ExportRTF

            Case "rbExportText"
                iExport = ExportType.ExportText

            Case "rbExportTIFF"
                iExport = ExportType.ExportTIFF

        End Select

    End Sub


    Private Sub frmMainMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub cmdStyle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubReport.Click

    End Sub

    Private Sub cmdHyperLink_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHyperLink.Click

    End Sub

    Private Sub cmdInstalledPrinters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInstalledPrinters.Click
        Dim objMsg As New System.Text.StringBuilder()
        Dim cPrinter As String

        For Each cPrinter In PrinterSettings.InstalledPrinters
            objMsg.Append(cPrinter)
            objMsg.Append(ControlChars.CrLf)
        Next

        MsgBox(objMsg.ToString, MsgBoxStyle.OKOnly, "Installed Printers")

    End Sub
End Class
