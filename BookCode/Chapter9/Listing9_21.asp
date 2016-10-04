Function ExportReport() As Long
    ' Sample WebCache Report

    Dim rpt         As New ActiveReport
    Dim oWebCache   As New WebCache
    Dim oPDF        As New ActiveReportsPDFExport.ARExportPDF
    Dim aByteArray  As Variant
    Dim lReturnID   As Long
    
    ' Load the Report
    rpt.LoadLayout ("\Test.RPX")
    
    ' Run the Report
    rpt.Run
    
    ' Export the Report to a Byte Array
    Call oPDF.ExportStream(rpt.Pages, aByteArray)
    
    ' Now Cache the Byte Array
    lReturnID = oWebCache.CacheContent("Application/PDF", aByteArray)
    
    ' Return the Cached Report ID, which is how you can later
    ' retrieve this Cached report document
    ExportReport = lReturnID
End Function
