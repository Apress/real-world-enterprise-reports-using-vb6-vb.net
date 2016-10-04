Option Explicit

Dim objVSPrinter1 As New VSPrinter
Dim objVSPDF As New VSPDF
    
Dim bIsWebReport As Boolean
Dim lUserID As Long

Public Property Get UserID() As Variant
    UserID = lUserID
End Property

Public Property Let UserID(ByVal vNewValue As Variant)
    lUserID = vNewValue
End Property

Public Property Get IsWebReport() As Variant
    IsWebReport = bIsWebReport
End Property

Public Property Let IsWebReport(ByVal vNewValue As Variant)
    bIsWebReport = vNewValue
End Property

Public Function MyFirstWebReport() As String
    Dim cFileName As String
    Dim x As Integer
    
    cFileName = "c:\temp\" & GetTempFileName(lUserID, "pdf")
        
    With objVSPrinter1
            
        .StartDoc
        
        .Preview = True
        
        .StartTable
                
        .TableCell(tcCols) = 2
        .TableCell(tcColWidth, , 1) = 3900
        .TableCell(tcColWidth, , 2) = 3900
        
        For x = 1 To 1000
        
            .TableCell(tcInsertRow) = x
            .TableCell(tcText, x, 1) = "Column 1 - Data Element " & x
            .TableCell(tcText, x, 2) = "Column 2 - Data Element " & x
            
        Next x
        
        .EndTable
        
        .EndDoc
        
        If bIsWebReport Then
            objVSPDF.ConvertDocument objVSPrinter1, cFileName
        End If
        
    End With
    
    Set objVSPrinter1 = Nothing
    Set objVSPDF = Nothing
    
    MyFirstWebReport = cFileName
    
End Function

Private Sub VSPrinter1_NewPage()
    
    objVSPrinter1.AddTable "3900|3900", _
        "Header 1|Header 2", vbNullString

End Sub

Function GetTempFileName(lUserID As Long, _
    cExtension As String) As String
    Dim cResult As String
    
    cResult = Month(Now) & _
              Day(Now) & _
              Year(Now) & _
              Hour(Now) & _
              Minute(Now) & _
              Second(Now) & "_" & _
              lUserID & "." & cExtension
    
    GetTempFileName = cResult

End Function
