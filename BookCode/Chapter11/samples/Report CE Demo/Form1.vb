Public Class Form1
    Inherits System.Windows.Forms.Form
    Friend WithEvents btnRunReport As System.Windows.Forms.Button
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu

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
        Me.btnRunReport = New System.Windows.Forms.Button
        '
        'btnRunReport
        '
        Me.btnRunReport.Location = New System.Drawing.Point(120, 244)
        Me.btnRunReport.Size = New System.Drawing.Size(112, 20)
        Me.btnRunReport.Text = "Run Report"
        '
        'Form1
        '
        Me.Controls.Add(Me.btnRunReport)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMenu1
        Me.MinimizeBox = False
        Me.Text = "Report CE Demo"

    End Sub

#End Region

  Private Sub btnRunReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRunReport.Click
    Dim pi As ProcessInfo
    Dim si() As Byte
    Dim intResult As Int32

  ' Launch the Report CE engine passing to it the sample report.
    intResult = LaunchApplication("\Program Files\Report CE\ReportCE.exe", "\Windows\Start Menu\Programs\Sample.rce", Nothing, Nothing, 0, _
      0, Nothing, Nothing, si, pi)

  End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
