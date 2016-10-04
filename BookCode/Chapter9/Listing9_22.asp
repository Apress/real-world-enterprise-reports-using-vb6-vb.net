Dim objWebReport As New WebReport()
Dim objPDF As New Pdf.PdfExport()
Dim objMemoryStream As System.IO.MemoryStream = _
    New System.IO.MemoryStream()

objWebReport.Run()

objPDF.Export(objWebReport.Document, objMemoryStream)

Response.BinaryWrite (objMemoryStream.ToArray())

Response.End()
