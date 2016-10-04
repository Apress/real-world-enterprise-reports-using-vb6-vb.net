Public Class Form1
    Inherits System.Windows.Forms.Form
    Friend WithEvents btnDisplayReport As System.Windows.Forms.Button
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu

    Dim cn As System.Data.SqlServerCe.SqlCeConnection
    Dim cmd As New System.Data.SqlServerCe.SqlCeCommand
    Dim dr As System.Data.SqlServerCe.SqlCeDataReader

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        MyBase.Dispose(disposing)
    End Sub

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Private Sub InitializeComponent()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.btnDisplayReport = New System.Windows.Forms.Button
        '
        'btnDisplayReport
        '
        Me.btnDisplayReport.Location = New System.Drawing.Point(116, 240)
        Me.btnDisplayReport.Size = New System.Drawing.Size(116, 24)
        Me.btnDisplayReport.Text = "Display Report"
        '
        'Form1
        '
        Me.Controls.Add(Me.btnDisplayReport)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMenu1
        Me.MinimizeBox = False
        Me.Text = "HTML Reporting"

    End Sub

#End Region

#Region " Event Procedures "

  Private Sub btnDisplayReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisplayReport.Click
    Dim ReportFile As String = "\Windows\Start Menu\Programs\Apress\report.html"
    Dim ReportHtml As String

  ' Retrieve the data.
    RetrieveData()

  ' Generate the report.
    ReportHtml = ProduceReport()

  ' Save the report.
    SaveReport(ReportHtml, ReportFile)

  ' Display the report.
    DisplayReport(ReportFile)

  End Sub

#End Region

#Region " General Procedures "

  Sub DisplayReport(ByVal ReportFile As String)
    Dim pi As ProcessInfo
    Dim si() As Byte
    Dim intResult As Int32

  ' Launch Pocket IE to display the report.
    intResult = LaunchApplication("\Windows\iexplore.exe", ReportFile, Nothing, Nothing, 0, _
      0, Nothing, Nothing, si, pi)

  End Sub

  Function ProduceReport() As String
    Dim HTML As String

  ' Build the report header.
    HTML = "<HTML>"
    HTML += "<BODY>"
    HTML += "<H1><FONT COLOR=Blue>Customer List</FONT></H1>"
    HTML += "<TABLE>"
    HTML += "<TR>"
    HTML += "<TD><B>COMPANY</B></TD>"
    HTML += "<TD><B>CONTACT</B></TD>"
    HTML += "<TD><B>PHONE</B></TD>"

  ' Loop through the data.
    While dr.Read()
      HTML += "<TR>"
      HTML += "<TD><FONT SIZE=-2>" & dr("CompanyName") & "</TD>"
      HTML += "<TD><FONT SIZE=-2>" & dr("ContactName") & "</TD>"
      HTML += "<TD><FONT SIZE=-2>" & dr("Phone") & "</TD>"
      HTML += "</TR>"
    End While

  ' Add the report footer.
    HTML += "</TABLE>"
    HTML += "</BODY>"
    HTML += "</HTML>"

  ' Return the report.
    Return HTML

  End Function

  Sub RetrieveData()

        ' Open the connection.
        Try
            cn = New _
              System.Data.SqlServerCe.SqlCeConnection( _
              "Data Source=\Windows\Start Menu\Programs\Apress\NorthwindDemo.sdf")
            cn.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ' Configure and execute the command.
        cmd.CommandText = "SELECT * FROM Customers"
        cmd.Connection = cn
        dr = cmd.ExecuteReader

  End Sub

  Sub SaveReport(ByVal ReportHtml As String, ByVal ReportFile As String)
    Dim sw As System.IO.StreamWriter

  ' Open the file.
    sw = New System.IO.StreamWriter(ReportFile)

  ' Write the report.
    sw.Write(ReportHtml)

  ' Close the file.
    sw.Close()

  End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
