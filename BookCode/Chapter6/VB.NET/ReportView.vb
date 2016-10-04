Imports System.Data.SqlClient

Public Class frmReportViewer
    Inherits System.Windows.Forms.Form

    Dim oReport As DataDynamics.ActiveReports.ActiveReport

    Property Report()
        Get
            Report = oReport
        End Get
        Set(ByVal Value)
            oReport = Value
        End Set
    End Property

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
    Friend WithEvents Viewer1 As DataDynamics.ActiveReports.Viewer.Viewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Viewer1 = New DataDynamics.ActiveReports.Viewer.Viewer()
        Me.SuspendLayout()
        '
        'Viewer1
        '
        Me.Viewer1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.Viewer1.BackColor = System.Drawing.SystemColors.Control
        Me.Viewer1.Location = New System.Drawing.Point(8, 8)
        Me.Viewer1.Name = "Viewer1"
        Me.Viewer1.ReportViewer.CurrentPage = 0
        Me.Viewer1.ReportViewer.MultiplePageCols = 3
        Me.Viewer1.ReportViewer.MultiplePageRows = 2
        Me.Viewer1.Size = New System.Drawing.Size(816, 504)
        Me.Viewer1.TabIndex = 0
        Me.Viewer1.TableOfContents.Text = "Contents"
        Me.Viewer1.TableOfContents.Width = 200
        Me.Viewer1.Toolbar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmReportViewer
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(832, 517)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Viewer1})
        Me.Name = "frmReportViewer"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Bookmark Report"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmReportViewer_Load(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles MyBase.Load

        Viewer1.Document = oReport.Document
        oReport.Run()

    End Sub

    Private Sub Viewer1_HyperLink(ByVal sender As Object, _
            ByVal e As DataDynamics.ActiveReports.Viewer.HyperLinkEventArgs) _
            Handles Viewer1.HyperLink

        Dim objSqlConnection As New SqlConnection()
        Dim objCommand As SqlCommand
        Dim objDR As SqlDataReader
        Dim cSQL As String

        objSqlConnection.ConnectionString = "Data Source=(Local);" & _
                                            "Initial catalog=Northwind;" & _
                                            "Integrated security=SSPI;" & _
                                            "Persist security info=False"
        objSqlConnection.Open()

        cSQL = "SELECT OrderDate, ShippedDate, Freight " & _
               "FROM Orders " & _
               "WHERE CustomerID = " & Chr(39) & e.HyperLink.ToString & Chr(39)

        objCommand = New SqlCommand(cSQL, objSqlConnection)

        objDR = objCommand.ExecuteReader()

        Dim objReport As New DrillDownDetail()

        objReport.DataSource = objDR

        Viewer1.Document = objReport.Document

        objReport.Run()

    End Sub

End Class
