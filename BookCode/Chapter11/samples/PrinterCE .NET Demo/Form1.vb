Imports FieldSoftware.PrinterCE_NetCF

Public Class Form1
    Inherits System.Windows.Forms.Form
    Friend WithEvents btnProduceReport As System.Windows.Forms.Button
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
Me.btnProduceReport = New System.Windows.Forms.Button
'
'btnProduceReport
'
Me.btnProduceReport.Location = New System.Drawing.Point(120, 244)
Me.btnProduceReport.Size = New System.Drawing.Size(112, 20)
Me.btnProduceReport.Text = "Produce Report"
'
'Form1
'
Me.Controls.Add(Me.btnProduceReport)
Me.MaximizeBox = False
Me.Menu = Me.MainMenu1
Me.MinimizeBox = False
Me.Text = "PrinterCE .NET Demo"

    End Sub

#End Region

  Private Sub btnProduceReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProduceReport.Click
    Dim prce As PrinterCE

    Try
  ' Create an instance of the PrinterCE component.
      prce = New PrinterCE(PrinterCE.EXCEPTION_LEVEL.ABORT_JOB, "YourLicense")

  ' Prompt the user for the target printer.
      prce.SelectPrinter(True)

  ' Print out a simple message.
      prce.DrawText("Hello World")

  ' Complete the print document, which in turn submits the print job.
      prce.EndDoc()

  ' Handle any errors that occur.
    Catch exc As PrinterCEException
      MessageBox.Show("PrinterCE Exception", "Exception")

  ' Clean up.
    Finally
      prce.ShutDown()
    End Try

  End Sub
End Class
