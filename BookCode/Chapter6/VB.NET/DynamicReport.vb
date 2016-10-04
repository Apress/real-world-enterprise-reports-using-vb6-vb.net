Imports DataDynamics.ActiveReports
Imports System.Data.SqlClient

Public Class frmDynamicReport
    Inherits System.Windows.Forms.Form

    Dim WithEvents objAR As New ActiveReport()
    Dim objDR As SqlDataReader
    Dim objCommand As SqlCommand

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
        Me.Viewer1.Size = New System.Drawing.Size(728, 352)
        Me.Viewer1.TabIndex = 0
        Me.Viewer1.TableOfContents.Text = "Contents"
        Me.Viewer1.TableOfContents.Width = 200
        Me.Viewer1.Toolbar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmDynamicReport
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(744, 365)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Viewer1})
        Me.Name = "frmDynamicReport"
        Me.Text = "Dynamic Report"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form3_Load(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles MyBase.Load

        'Call ColumnarRpt()

        Call Labels()

    End Sub

    Private Sub Labels_DataInitialize(ByVal sender As Object, ByVal e As System.EventArgs) _
        Handles objAR.DataInitialize

        objAR.Fields.Add("CSZ")

    End Sub

    Private Sub Labels_FetchData(ByVal sender As Object, _
        ByVal eArgs As DataDynamics.ActiveReports.ActiveReport.FetchEventArgs) _
        Handles objAR.FetchData

        If eArgs.EOF Then
            Exit Sub
        End If

        With objDR
            objAR.Fields("CSZ").Value = .Item("City") & ", " & _
                                        .Item("Region") & " " & _
                                        .Item("PostalCode")
        End With

    End Sub

    Sub Labels()
        Dim objDetail As DataDynamics.ActiveReports.Detail
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


        'Next, Detail section of the report object. No other
        'sections are needed
        objDetail = objAR.Sections.Add(SectionType.Detail, "Detail")

        'Two columns of labels that snake down and across the page
        With objDetail
            .ColumnCount = 2
            .ColumnDirection = ColumnDirection.DownAcross
            .KeepTogether = True
        End With

        '...and set the height of the detail section to one inch
        objAR.Sections(0).Height = 1


        'Create TextBox objects for the data columns and add 
        'them to section 0, the Details
        objTextBox = New TextBox()

        With objTextBox
            .Name = "fldCompanyName"
            .DataField = "CompanyName"
            .Width = 3
            .Left = 0.2
            .Height = 0.1
            .Top = 0
        End With

        objAR.Sections(0).Controls.Add(objTextBox)


        objTextBox = New TextBox()

        With objTextBox
            .Name = "fldContactName"
            .DataField = "ContactName"
            .Width = 3
            .Left = 0.2
            .Height = 0.1
            .Top = 0.13
        End With

        objAR.Sections(0).Controls.Add(objTextBox)


        objTextBox = New TextBox()

        With objTextBox
            .Name = "fldAddress"
            .DataField = "Address"
            .Width = 3
            .Left = 0.2
            .Height = 0.1
            .Top = 0.26
        End With

        objAR.Sections(0).Controls.Add(objTextBox)


        objTextBox = New TextBox()

        With objTextBox
            .Name = "fldCSZ"
            .DataField = "CSZ"
            .Width = 3
            .Left = 0.2
            .Height = 0.1
            .Top = 0.39
        End With

        objAR.Sections(0).Controls.Add(objTextBox)

        objAR.DataSource = objDR

        Viewer1.Document = objAR.Document

        objAR.Run()

    End Sub

    Sub ColumnarRpt()
        Dim objAR As New ActiveReport()
        Dim objSqlConnection As New SqlConnection()
        Dim objDR As SqlDataReader
        Dim objCommand As SqlCommand
        Dim cSQL As String
        Dim objLabel As Label
        Dim objTextBox As TextBox
        Dim objSection As Section

        'First, establish a data connection
        objSqlConnection.ConnectionString = "Data Source=(local);" & _
                                            "Initial catalog=Northwind;" & _
                                            "Integrated security=SSPI;" & _
                                            "Persist security info=False"
        objSqlConnection.Open()

        cSQL = "SELECT CompanyName, ContactName " & _
               "FROM Customers "

        objCommand = New SqlCommand(cSQL, objSqlConnection)

        objDR = objCommand.ExecuteReader()


        'Next, create the various sections of the report object
        With objAR.Sections
            .Add(SectionType.ReportHeader, "ReportHeader")
            .Add(SectionType.PageHeader, "PageHeader")
            .Add(SectionType.Detail, "Detail")
            .Add(SectionType.PageFooter, "PageFooter")
            .Add(SectionType.ReportFooter, "ReportFooter")
        End With

        '...and set the height of those sections measured in inches
        objAR.Sections(0).Height = 0.5
        objAR.Sections(1).Height = 0.3
        objAR.Sections(2).Height = 0.3
        objAR.Sections(3).Height = 0.3
        objAR.Sections(4).Height = 0.5


        ' CType(objAR.Sections.Add(SectionType.GroupHeader, "GH1"), GroupHeader).DataField() = "CompanyName"

        'Create label objects for the column headers and add 
        'the to section 1, the PageHeader
        objLabel = New Label()

        With objLabel
            .Name = "Label1"
            .Text = "Company Name"
            .Width = 3
            .Left = 0.5
            .Height = 0.3
        End With

        objAR.Sections(1).Controls.Add(objLabel)


        objLabel = New Label()

        With objLabel
            .Name = "Label2"
            .Text = "Contact Name"
            .Width = 3
            .Left = 3.5
            .Height = 0.3
        End With

        objAR.Sections(1).Controls.Add(objLabel)



        'Create TextBox objects for the data columns and add 
        'the to section 2, the Details
        objTextBox = New TextBox()

        With objTextBox
            .Name = "TextBox1"
            .DataField = "CompanyName"
            .Width = 3
            .Left = 0.5
            .Height = 0.3
        End With

        objAR.Sections(2).Controls.Add(objTextBox)


        objTextBox = New TextBox()

        With objTextBox
            .Name = "TextBox2"
            .DataField = "ContactName"
            .Width = 3
            .Left = 3.5
            .Height = 0.3
        End With

        objAR.Sections(2).Controls.Add(objTextBox)


        'Finally, set the data source of the report
        objAR.DataSource = objDR

        '...set the document property of the viewer to
        'that of the report object
        Viewer1.Document = objAR.Document

        '...and away you go
        objAR.Run()

    End Sub

    Private Sub Viewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Viewer1.Load

    End Sub
End Class
