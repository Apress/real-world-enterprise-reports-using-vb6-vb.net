Private Sub Button1_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) _
    Handles Button1.Click

    Dim objRunReport As New localhost1.Service1()
    Dim cWhere As String
    Dim cEncodedFile As String

    Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

    cWhere = "State = 'NJ'"

    objRunReport.EmployeeRpt(cWhere, localhost1.ExportFormats.PDF)

    cEncodedFile = objRunReport.SendReportToUser("c:\docs\reports\chapter4.pdf")

    Call Base64Decode(cEncodedFile, "c:\docs\myreportdoc.pdf")

    Cursor.Current = System.Windows.Forms.Cursors.Default

End Sub

Sub Base64Decode(ByVal cData As String, ByVal cFileName As String)
    Dim aData() As Byte
    Dim objFileStream As System.IO.FileStream

    aData = System.Convert.FromBase64String(cData)

    objFileStream = New System.IO.FileStream(cFileName, _
        IO.FileMode.Create, IO.FileAccess.Write)

    objFileStream.Write(aData, 0, aData.Length - 1)

    objFileStream.Close()

End Sub
