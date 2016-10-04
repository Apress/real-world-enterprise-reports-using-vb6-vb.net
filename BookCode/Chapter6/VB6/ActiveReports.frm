VERSION 5.00
Begin VB.Form frmMainMenu 
   Caption         =   "ActiveReports Demo"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReport 
      Caption         =   "Report Viewer"
      Height          =   495
      Index           =   14
      Left            =   3720
      TabIndex        =   21
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "SubReport"
      Height          =   495
      Index           =   13
      Left            =   3720
      TabIndex        =   20
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Mail Merge 2"
      Height          =   495
      Index           =   12
      Left            =   1920
      TabIndex        =   19
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "XML"
      Height          =   495
      Index           =   11
      Left            =   1920
      TabIndex        =   18
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Mail Merge"
      Height          =   495
      Index           =   10
      Left            =   1920
      TabIndex        =   17
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Drill Down"
      Height          =   495
      Index           =   9
      Left            =   1920
      TabIndex        =   16
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Unbound Data"
      Height          =   495
      Index           =   8
      Left            =   1920
      TabIndex        =   15
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Designer"
      Height          =   495
      Index           =   7
      Left            =   1920
      TabIndex        =   14
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "SpreadBuilder"
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Export Options"
      Height          =   1695
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton cmdReport 
         Caption         =   "Export Report"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1935
      End
      Begin VB.OptionButton rbExport 
         Caption         =   "RTF"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton rbExport 
         Caption         =   "HTML"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton rbExport 
         Caption         =   "TIFF"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton rbExport 
         Caption         =   "Text"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton rbExport 
         Caption         =   "PDF"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton rbExport 
         Caption         =   "Excel"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "From Scratch"
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Grouping"
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Get Printer List"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Labels"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "First Report"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum ExportFormat

    Excel = 0
    PDF = 1
    Text = 2
    TIFF = 3
    HTML = 4
    RTF = 5
    
End Enum


Private Sub cmdReport_Click(Index As Integer)

    Select Case Index
    
        Case 0
            Call RunFirstReport
            
        Case 1
            Call RunLabel
    
        Case 2
            Call GetPrinterList
    
        Case 3
            Call RunGrouping
            
        Case 4
            Call RunFromScratch
            
        Case 5
            Call Export
            
        Case 6
            Call SpreadBuilder
            
        Case 7
            Call Designer
            
        Case 8
            Call RunUnboundData
            
        Case 9
            Call RunDrillDown
            
        Case 10
            Call RunMailMerge
        
        Case 11
            Call RunXML
            
        Case 12
            Call RunMailMerge2
            
        Case 13
            Call RunSubReport
            
        Case 14
            Call Viewer
            
    End Select

End Sub

Sub RunFirstReport()
    
    FirstReport.Show

End Sub

Sub RunLabel()

    Labels.Show

End Sub

Sub RunGrouping()

    Grouping.Show

End Sub

Sub RunFromScratch()

    FromScratch.Show

End Sub

Sub RunDrillDown()

    DrillDown.Show

End Sub

Sub RunMailMerge()

    MailMerge.Show
    
End Sub

Sub RunSubReport()

    SubReport.Show
    
End Sub

Sub GetPrinterList()
    Dim x As Integer
    Dim cMsg As String
    
    With Labels.Printer
    
        For x = 0 To .NDevices - 1
            cMsg = cMsg & .Devices(x) & vbCrLf
        Next
    
        cMsg = cMsg & vbCrLf
        
        cMsg = cMsg & "Current device: " & .DeviceName & vbCrLf
        
        cMsg = cMsg & "Current port: " & .Port & vbCrLf
        
        cMsg = cMsg & vbCrLf
        
        For x = 0 To UBound(.PaperBinNames) - 1
            cMsg = cMsg & .PaperBinNames(x) & vbCrLf
        Next
    
    End With
    
    MsgBox cMsg
    
End Sub

Sub Export()
    Dim objExcelExport As ActiveReportsExcelExport.ARExportExcel
    Dim objPDFExport As ActiveReportsPDFExport.ARExportPDF
    Dim objTextExport As ActiveReportsTextExport.ARExportText
    Dim objTIFFExport As ActiveReportsTIFFExport.TIFFExport
    Dim objHTMLExport As ActiveReportsHTMLExport.HTMLexport
    Dim objRTFExport As ActiveReportsRTFExport.ARExportRTF
    Dim iExport As Integer
    
    
    For iExport = 0 To rbExport.Count - 1
        
        If rbExport(iExport) Then
            Exit For
        End If
        
    Next iExport
    
    
    Grouping.Run
    
    Select Case iExport
    
        Case ExportFormat.Excel
            Set objExcelExport = New ActiveReportsExcelExport.ARExportExcel
            With objExcelExport
               .FileName = App.Path & "\Grouping.xls"
               .MinColumnWidth = 2000
               .GenPagebreaks = True
               .MultiSheet = True
               .Version = 8
               .Export Grouping.Pages
            End With
                    
        Case ExportFormat.PDF
            Set objPDFExport = New ActiveReportsPDFExport.ARExportPDF
            objPDFExport.FileName = App.Path & "\Grouping.PDF"
            objPDFExport.Export Grouping.Pages
        
        Case ExportFormat.Text
            Set objTextExport = New ActiveReportsTextExport.ARExportText
            objTextExport.FileName = App.Path & "\Grouping.txt"
            objTextExport.TextDelimiter = ","
            objTextExport.Export Grouping.Pages
            
        Case ExportFormat.TIFF
            Set objTIFFExport = New ActiveReportsTIFFExport.TIFFExport
            objTIFFExport.FileName = App.Path & "\Grouping.tif"
            objTIFFExport.Export Grouping.Pages
            
        Case ExportFormat.HTML
            Set objHTMLExport = New ActiveReportsHTMLExport.HTMLexport
            objHTMLExport.FileName = App.Path & "\Grouping.html"
            objHTMLExport.Title = "Sample Report"
            objHTMLExport.Export Grouping.Pages
            
        Case ExportFormat.RTF
            Set objRTFExport = New ActiveReportsRTFExport.ARExportRTF
            objRTFExport.FileName = App.Path & "\Grouping.rtf"
            objRTFExport.Export Grouping.Pages
            
    End Select
    
    MsgBox "Export Complete"
    
End Sub

Sub SpreadBuilder()
    Dim objSpreadBuilder As New ActiveReportsExcelExport.SpreadBuilder
    Dim oConn As New ADODB.Connection
    Dim oRS As ADODB.Recordset
    Dim cSQL As String
    Dim x As Integer
    
    oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & GetPath & "SampleDatabase.mdb;" & _
    "Persist Security Info=False"
    
    oConn.Open
    
    cSQL = "SELECT * FROM Requester"
    
    Set oRS = oConn.Execute(cSQL)
    
    objSpreadBuilder.Sheets.Add "Sheet1"
    
    
    With objSpreadBuilder.Sheets(0)
    
        .Rows(0).Height = 800
    
        For x = 0 To 7
            .Columns(x).Width = 1600
            
            .Cell(0, x).FontBold = True
            .Cell(0, x).ForeColor = vbBlue
        Next x
        
        x = 0
        
        .Cell(x, 0).Value = "Last Name"
        .Cell(x, 1).Value = "First Name"
        .Cell(x, 2).Value = "Address 1"
        .Cell(x, 3).Value = "Address 2"
        .Cell(x, 4).Value = "City"
        .Cell(x, 5).Value = "State"
        .Cell(x, 6).Value = "Zip"
        
        .Cell(x, 7).Value = "Amount"
        .Cell(x, 7).Alignment = SBAlignRight
        
        x = 1
    
        Do While Not oRS.EOF
                        
            .Cell(x, 0).Value = "" & oRS("LastName")
            .Cell(x, 1).Value = "" & oRS("FirstName")
            .Cell(x, 2).Value = "" & oRS("Address1")
            .Cell(x, 3).Value = "" & oRS("Address2")
            .Cell(x, 4).Value = "" & oRS("City")
            .Cell(x, 5).Value = "" & oRS("State")
            .Cell(x, 6).Value = "" & oRS("Zip")
            
            .Cell(x, 7).Type = SBNumber
            
            If IsNull(oRS("Amount")) Then
                .Cell(x, 7).Value = 0
                .Cell(x, 7).NumberFormat = "###,##0.00"
            Else
                .Cell(x, 7).Value = oRS("Amount")
            End If
    
            
            x = x + 1
             
            oRS.MoveNext
         
        Loop
     
    End With
    
    oRS.Close
    Set oRS = Nothing
     
    objSpreadBuilder.Save App.Path & "\myfile.xls"
    
    Set objSpreadBuilder = Nothing
    
    MsgBox "Saved to " & App.Path & "\myfile.xls"
    
End Sub

Sub Designer()
    frmDesigner.Show vbModal
End Sub

Sub RunUnboundData()
    UnboundData.Show
End Sub

Sub RunXML()
    XML.Show
End Sub

Sub RunMailMerge2()
    MailMerge2.Show
End Sub

Sub Viewer()
    frmARViewer.Show
End Sub

Function GetPath() As String
    Dim cResult As String
    Dim iPos As Integer
    
    cResult = App.Path
    
    iPos = InStrRev(cResult, "\")
    
    If iPos > 0 Then
    
        cResult = Mid(cResult, 1, iPos)
        
    End If

    GetPath = cResult

End Function

