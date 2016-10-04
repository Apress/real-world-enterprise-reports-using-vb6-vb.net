VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{1BCC7098-34C1-4749-B1A3-6C109878B38F}#1.0#0"; "vspdf8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Begin VB.Form Form1 
   Caption         =   "VS-View demo"
   ClientHeight    =   11865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   ScaleHeight     =   11865
   ScaleWidth      =   14355
   StartUpPosition =   2  'CenterScreen
   Begin VSPrinter8LibCtl.VSPrinter VSPrinter1 
      Height          =   11655
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   12015
      _cx             =   21193
      _cy             =   20558
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   68.4659090909091
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Hyperlinks"
      Height          =   372
      Index           =   16
      Left            =   12240
      TabIndex        =   20
      Top             =   7800
      Width           =   2052
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Export to PDF"
      Height          =   372
      Index           =   15
      Left            =   12240
      TabIndex        =   19
      Top             =   7320
      Width           =   2052
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   12360
      Top             =   9600
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPivotTable 
      Caption         =   "Pivot Table"
      Height          =   375
      Left            =   12240
      TabIndex        =   18
      Top             =   9000
      Width           =   2055
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Get Printer List"
      Height          =   375
      Index           =   14
      Left            =   12240
      TabIndex        =   17
      Top             =   6840
      Width           =   2055
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Cross Tab SQL"
      Height          =   375
      Index           =   13
      Left            =   12240
      TabIndex        =   16
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Cross Tab"
      Height          =   375
      Index           =   12
      Left            =   12240
      TabIndex        =   15
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Signature Display"
      Height          =   375
      Index           =   11
      Left            =   12240
      TabIndex        =   14
      Top             =   5400
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   12240
      Picture         =   "VSView.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   13
      Top             =   8400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Array-based Report"
      Height          =   375
      Index           =   10
      Left            =   12240
      TabIndex        =   12
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Table Dump"
      Height          =   375
      Index           =   9
      Left            =   12240
      TabIndex        =   11
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Snaked Column"
      Height          =   375
      Index           =   8
      Left            =   12240
      TabIndex        =   10
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Free Form Text - RTF"
      Height          =   375
      Index           =   7
      Left            =   12240
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Free Form Text - Plain"
      Height          =   375
      Index           =   6
      Left            =   12240
      TabIndex        =   8
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Labels"
      Height          =   375
      Index           =   5
      Left            =   12240
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Standard Columnar Report"
      Height          =   375
      Index           =   4
      Left            =   12240
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Header/Footer"
      Height          =   375
      Index           =   3
      Left            =   12240
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Styles"
      Height          =   375
      Index           =   2
      Left            =   12240
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Address Report"
      Height          =   375
      Index           =   1
      Left            =   12240
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Sample Report"
      Height          =   375
      Index           =   0
      Left            =   12240
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdSelectPrinter 
      Caption         =   "Select Printer"
      Height          =   375
      Left            =   12240
      TabIndex        =   1
      Top             =   11400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox lstPrinters 
      Height          =   1035
      Left            =   12240
      TabIndex        =   0
      Top             =   10080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VSPDF8LibCtl.VSPDF8 VSPDF81 
      Left            =   13920
      Top             =   8520
      Author          =   ""
      Creator         =   ""
      Title           =   ""
      Subject         =   ""
      Keywords        =   ""
      Compress        =   3
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cHdrFontName As String
Dim bHdrFontBold As Boolean
Dim iHdrFontSize As Integer
Dim iReport As Integer

Dim oConn As New ADODB.Connection

Public Enum DumpType
    dtASCII = 0
    dtExcel = 1
End Enum

Public Enum LetterType
    ltPlainText = 1
    ltRTF = 2
End Enum

Public Enum Report
    Sample = 0
    Address = 1
    Style = 2
    HeaderFooter = 3
    StandardColumnar = 4
    Label = 5
    TextPlainText = 6
    TextRTF = 7
    SnakedColumn = 8
    DumpTabledata = 9
    ArrayBased = 10
    SignatureDisplaygraphic = 11
    CrossTab = 12
    CrossTabSQLSELECT = 13
    GetPrinterListBox = 14
    ExportToPDF = 15
    Hyperlinks = 16
End Enum

Private Sub cmdPivotTable_Click()
    Dim oRS As ADODB.Recordset
    Dim objExcel As Excel.Application
    Dim objWorkBook As Excel.Workbook
    Dim objWS As Excel.Worksheet
    Dim objPivotWS As Excel.Worksheet
    Dim iRow As Integer
    Dim cSQL As String
    Dim cRange As String
    
    cSQL = "SELECT s.Last & ', ' & s.First AS name, p.Title, t.Score " & _
            "FROM Student s, Program p, Test t " & _
            "WHERE s.ID = t.Studnum " & _
            "AND p.ID = t.TestID"
    
    Set oRS = oConn.Execute(cSQL)
    
        
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    
    objExcel.ScreenUpdating = False
    
    Set objWorkBook = objExcel.Workbooks.Add
    Set objWS = objExcel.Worksheets.Add
    
    iRow = 1
        
    objWS.Cells(iRow, 1) = "Student"
    objWS.Cells(iRow, 2) = "Class"
    objWS.Cells(iRow, 3) = "Average"
    
    iRow = 2
            
            
    Do While Not oRS.EOF
        
        objWS.Cells(iRow, 1) = oRS("name")
        objWS.Cells(iRow, 2) = oRS("Title")
        objWS.Cells(iRow, 3) = oRS("Score")
        
        iRow = iRow + 1
            
        oRS.MoveNext
        
    Loop
    
    iRow = iRow - 1
    
    cRange = "Sheet4!R1C1:R" & iRow & "C3"
        
    objWorkBook.PivotCaches.Add(xlDatabase, cRange).CreatePivotTable "", _
        "MyPivot", xlPivotTableVersion10
    
    Set objPivotWS = objExcel.ActiveSheet
    
    With objPivotWS.PivotTables("MyPivot").PivotFields("Student")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With objPivotWS.PivotTables("MyPivot").PivotFields("Class")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    With objPivotWS.PivotTables("MyPivot")
        .AddDataField objPivotWS.PivotTables("MyPivot").PivotFields("Average"), _
            vbNullString, xlAverage
        .ColumnGrand = True
        .RowGrand = True
        .GrandTotalName = "Subject Averages"
        .EnableDrilldown = True
    End With
    
    objWorkBook.ShowPivotTableFieldList = False
    
    objExcel.Application.CommandBars("PivotTable").Enabled = False
     
    objExcel.ScreenUpdating = True
     
End Sub

Private Sub cmdReport_Click(Index As Integer)
    
    iReport = Index
    
    If iReport = Report.GetPrinterListBox Then
        lstPrinters.Visible = True
        cmdSelectPrinter.Visible = True
    Else
        lstPrinters.Visible = False
        cmdSelectPrinter.Visible = False
    End If
    
    With VSPrinter1
        .Clear
        .Header = vbNullString
        .Footer = vbNullString
    End With
    

    Select Case iReport
    
        Case Report.Sample
            Call SampleRpt
        
        Case Report.Address
            Call AddressRpt
        
        Case Report.Style
            Call Styles
        
        Case Report.HeaderFooter
            Call HeaderFooterRpt
        
        Case Report.StandardColumnar
            Call StandardColumnarRpt
        
        Case Report.Label
            Call Labels
        
        Case Report.TextPlainText
            Call FreeFormText(ltPlainText)
        
        Case Report.TextRTF
            Call FreeFormText(ltRTF)
        
        Case Report.SnakedColumn
            Call SnakedColumnRpt
        
        Case Report.DumpTabledata
            Call SampleRpt
            Call DumpTable(VSPrinter1, dtExcel)
        
        Case Report.ArrayBased
            Call ArrayBasedRpt
        
        Case Report.SignatureDisplaygraphic
            Call SignatureDisplay
    
        Case Report.CrossTab
            Call CrossTabRpt
            
        Case Report.CrossTabSQLSELECT
            Call CrossTabSQL
                                
        Case Report.GetPrinterListBox
            Call GetPrinterList
            
        Case Report.ExportToPDF
            Call ExportReportToPDF
            
        Case Report.Hyperlinks
            Call JumpHyperlinks
            
    End Select

End Sub

Private Sub cmdSelectPrinter_Click()
    VSPrinter1.Device = lstPrinters.Text
End Sub

Private Sub Form_Load()

    oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BookCode\SampleDatabase.mdb;Persist Security Info=False"
    oConn.Open

End Sub

Sub Styles()
    Dim cFontName As String
    Dim dblFontSize As Double
    Dim bFontBold As Boolean
    Dim iSpaceAfter As Integer
        
    With VSPrinter1
        cFontName = .FontName
        dblFontSize = .FontSize
        bFontBold = .FontBold
        iSpaceAfter = .SpaceAfter
    
        .FontName = "Arial"
        .FontSize = 48
        .FontBold = True
        .SpaceAfter = 2000
        .Styles.Add "Chapter Title", vpsAll
        
        .FontName = "Arial"
        .FontSize = 16
        .FontBold = False
        .LineSpacing = 250
        .Styles.Add "Intro Text", vpsAll
                
        .StartDoc
        
        .Styles.Apply "Chapter Title"
        .Paragraph = "Chapter 1"
        
        .Styles.Apply "Intro Text"
        .Paragraph = "It was the best of times, it was the worst of times. " & _
                    "It was the age of wisdom, it was the age of foolishness..."
                    
        .FontName = cFontName
        .FontSize = dblFontSize
        .FontBold = bFontBold
        .SpaceAfter = iSpaceAfter
        
        .EndDoc
    End With

End Sub

Sub GetPrinterList()
    Dim x As Integer
    
    For x = 0 To VSPrinter1.NDevices - 1
        lstPrinters.AddItem VSPrinter1.Devices(x)
    Next
    
End Sub

Sub DumpTable(oPrinter As VSPrinter, iDumpType As DumpType)
    Dim objExcel As Excel.Application
    Dim objWorkBook As Excel.Workbook
    Dim objWS As Excel.Worksheet
    Dim x As Integer
    Dim y As Integer
    Dim iRows As Integer
    Dim iCols As Integer
    Dim cLine As String
    Dim cFileName As String
    Dim cData As Variant
    
    Screen.MousePointer = vbHourglass
    
    If iDumpType = dtASCII Then
        cFileName = App.Path & "\ASCIIDump.txt"
        
        Open cFileName For Output As #1
    Else
        Set objExcel = CreateObject("Excel.Application")
        objExcel.Visible = False
        
        Set objWorkBook = objExcel.Workbooks.Add
        Set objWS = objExcel.Worksheets.Add
        
        cFileName = App.Path & "\ASCIIDump.xls"
        
    End If
    
    With VSPrinter1
    
        iRows = .TableCell(tcRows)
        iCols = .TableCell(tcCols)
                
        For x = 1 To iRows
        
            cLine = vbNullString
        
            For y = 1 To iCols
            
                cData = .TableCell(tcText, x, y)
            
                If iDumpType = dtASCII Then
                
                    cLine = cLine & Chr(34) & cData & Chr(34) & ","
                
                Else
                    
                    If Not IsNull(cData) Then
            
                        If IsNumeric(cData) And _
                           cData <> Empty And InStr(cData, ")") = 0 Then
                                            
                            objWS.Cells(x, y) = CDbl(cData)
                            
                        Else
                        
                            If IsNumeric(Mid(cData, 1, 1)) And _
                               Occurs(cData, "/") = 2 Then
                                
                                If IsDate(CStr("" & cData)) Then
                                    objWS.Cells(x, y) = Format(CStr(cData), "mm/dd/yyyy")
                                Else
                                    objWS.Cells(x, y) = Chr(39) & cData
                                End If
                                
                            Else
                            
                                objWS.Cells(x, y) = Chr(39) & cData
                                
                            End If
                        
                        End If
                                        
                    End If
                
                End If
                
            Next y
        
            If iDumpType = dtASCII Then
                Print #1, Mid$(cLine, 1, Len(cLine) - 1)
            End If
            
        Next x
    
    End With
        
    If iDumpType = dtASCII Then
        
        Close #1
    
    Else
    
        If Dir(cFileName) <> vbNullString Then
            Kill cFileName
        End If
        
        objWS.Cells.ColumnWidth = 40
        
        objWS.SaveAs cFileName
    
        objWorkBook.Close
                       
        objExcel.Quit
        
    End If
    
    MsgBox "Report has been exported to " & cFileName, vbOKOnly

    Screen.MousePointer = vbDefault
    
End Sub

Function Occurs(ByVal cString As String, vntSearch As Variant) As Integer
    Dim iOrigLen As Integer
    
    iOrigLen = Len(cString)
    
    cString = Replace(cString, vntSearch, "")
    
    Occurs = iOrigLen - Len(cString)
    
End Function

Sub SnakedColumnRpt()
    Dim oRS As ADODB.Recordset
    Dim cSQL As String
    Dim iCol As Integer
    Dim iCurrentY As Integer
    Dim iLeft As Integer
    Dim iMarginLeft As Integer
    Dim cMonth As String
    Dim cCoordinator As String
    Dim cDateRange As String
    Dim cConventionName As String
    Dim cCityState As String
    Dim cConvNumber As String
    
    cSQL = "SELECT ID, ConventionName, StartDate, " & _
           "EndDate, City, State, Coordinator " & _
           "FROM Convention " & _
           "ORDER BY StartDate"
            
    Set oRS = oConn.Execute(cSQL)
    
    If oRS.EOF Then
        Exit Sub
    End If
    
    iCol = 1
    iLeft = iMarginLeft
    
    With VSPrinter1
        .TableBorder = tbNone
        .PageBorder = pbNone
        .Orientation = orLandscape
        .FontName = "Arial"
        .FontSize = 12
        .MarginLeft = 500
        .StartDoc
        
        iCurrentY = .CurrentY
        iMarginLeft = .MarginLeft
        iLeft = .MarginLeft
        
        cMonth = Format(oRS("StartDate"), "mmmm")
                                
        Call PrintMonth(iLeft, cMonth)
                        
        Do While Not oRS.EOF
        
            cCoordinator = oRS("Coordinator")
            cDateRange = Format(oRS("StartDate"), "mm/dd")
            cConventionName = UCase$(oRS("ConventionName"))
            cCityState = oRS("City") & ", " & oRS("State")
            cConvNumber = "#" & oRS("id")
                
            With VSPrinter1
            
                .MarginLeft = iLeft
                        
                .StartTable
                
                .TableCell(tcFontSize) = 12
                .TableCell(tcFontBold) = True
            
                .TableCell(tcCols) = 2
                .TableCell(tcRows) = 4
                
                .TableCell(tcColWidth, 1, 1) = 2000
                .TableCell(tcColWidth, 1, 2) = 1500
                
                .TableCell(tcAlign, 1, 2) = taRightMiddle
                
                .TableCell(tcColWidth, 3, 1) = 2500
                .TableCell(tcColWidth, 3, 2) = 1000
                
                .TableCell(tcAlign, 3, 2) = taRightMiddle
                
                .TableCell(tcFontBold, 2, 1) = True
                .TableCell(tcColSpan, 2) = 2
                        
                .TableCell(tcText, 1, 1) = cCoordinator
                .TableCell(tcText, 1, 2) = cDateRange
                .TableCell(tcText, 2, 1) = cConventionName
                .TableCell(tcText, 3, 1) = cCityState
                .TableCell(tcText, 3, 2) = cConvNumber
                
                .EndTable
                        
            End With

        
            oRS.MoveNext
            
            If oRS.EOF Then
            
                Exit Do
            
            End If
            
            If Format(oRS("StartDate"), "mmmm") <> cMonth Then

                If .CurrentY > 9300 Then

                    If iCol < 4 Then
                        iCol = iCol + 1
                        iLeft = iMarginLeft + 3500 * (iCol - 1)
                        .CurrentY = iCurrentY

                    Else
                        iCol = 1
                        iLeft = iMarginLeft
                        .NewPage
                    End If

                End If

                cMonth = Format(oRS("StartDate"), "mmmm")

                 Call PrintMonth(iLeft, cMonth)

            End If

            If .CurrentY > 9300 Then
                           
                If iCol < 4 Then
                    iCol = iCol + 1
                    iLeft = iMarginLeft + 3500 * (iCol - 1)
                Else
                    iCol = 1
                    iLeft = iMarginLeft
                    .NewPage
                End If
                
                .CurrentY = iCurrentY
                
            End If
        
        Loop
        
        .EndDoc
               
        oRS.Close
        Set oRS = Nothing
        
    End With
    
End Sub

Sub PrintMonth(iLeft As Integer, cMonth As String)

    With VSPrinter1
    
        .MarginLeft = iLeft
    
        .TableBorder = tbAll
        
        .StartTable
        
        .TableCell(tcCols) = 1
        .TableCell(tcRows) = 1
        
        .TableCell(tcFontSize) = 14
        .TableCell(tcFontBold) = True
        
        .TableCell(tcColWidth, , 1) = 3500
        
        .TableCell(tcText, 1, 1) = cMonth
        
        .EndTable

        .TableBorder = tbNone

    End With
    
End Sub

Sub SnakedColumnRpt1()
    Dim oRS As ADODB.Recordset
    Dim cSQL As String
    Dim iCol As Integer
    Dim iCurrentY As Integer
    Dim iLeft As Integer
    Dim iMarginLeft As Integer
    Dim cMonth As String
        
    cSQL = "SELECT ID, ConventionName, StartDate, " & _
           "EndDate, City, State, Coordinator " & _
           "FROM Convention " & _
           "ORDER BY StartDate"
            
    Set oRS = oConn.Execute(cSQL)
    
    If oRS.EOF Then
        Exit Sub
    End If
    
    iCol = 1
    iLeft = iMarginLeft
    
    With VSPrinter1
        .TableBorder = tbNone
        .PageBorder = pbNone
        .Orientation = orLandscape
        .FontName = "Arial"
        .FontSize = 12
        .MarginLeft = 500
        .StartDoc
        
        iCurrentY = .CurrentY
        iMarginLeft = .MarginLeft
        iLeft = .MarginLeft
        
        cMonth = Format(oRS("StartDate"), "mmmm")
                                
        Call PrintMonth(iLeft, cMonth)
                        
        Do While Not oRS.EOF
        
            Call PrintConference(iLeft, oRS)
        
            oRS.MoveNext
            
            If oRS.EOF Then
            
                Exit Do
            
            End If
            
            If Format(oRS("StartDate"), "mmmm") <> cMonth Then
            
                If .CurrentY > 9300 Then
    
                    If iCol < 4 Then
                        iCol = iCol + 1
                        iLeft = iMarginLeft + 3500 * (iCol - 1)
                        .CurrentY = iCurrentY
    
                    Else
                        iCol = 1
                        iLeft = iMarginLeft
                        .NewPage
                    End If
    
                End If
                
                cMonth = Format(oRS("StartDate"), "mmmm")
            
                 Call PrintMonth(iLeft, cMonth)
                
            End If
        
            If .CurrentY > 9300 Then
                
                .CurrentY = iCurrentY
                           
                If iCol < 4 Then
                    iCol = iCol + 1
                    iLeft = iMarginLeft + 3500 * (iCol - 1)
                    
                Else
                    iCol = 1
                    iLeft = iMarginLeft
                    '.MarginLeft = iMarginLeft
                    .NewPage
                End If
                
            End If
        
        Loop
               
        oRS.Close
        Set oRS = Nothing
        
    End With
    
End Sub

Sub PrintMonth1(iLeft As Integer, cMonth As String)

    With VSPrinter1
    
        .MarginLeft = iLeft
    
        .TableBorder = tbAll
        
        .StartTable
        
        .TableCell(tcCols) = 1
        .TableCell(tcRows) = 1
        
        .TableCell(tcFontSize) = 14
        .TableCell(tcFontBold) = True
        
        .TableCell(tcColWidth, , 1) = 3500
        
        .TableCell(tcText, 1, 1) = cMonth
        
        .EndTable

        .TableBorder = tbNone

    End With
    
End Sub

Sub PrintConference(iLeft As Integer, oRS As ADODB.Recordset)

    

End Sub

Sub FreeFormText(iLetterType As LetterType)
    Dim oRS As ADODB.Recordset
    Dim cSQL As String
    Dim cName As String
    Dim cSalutation As String
    Dim cLastName As String
    Dim cAddress1 As String
    Dim cAddress2 As String
    Dim cCSZ As String
    Dim cCaseNumber As String
    Dim cLetterText As String
    Dim cText As String
    
            
    cSQL = "SELECT LetterText " & _
            "FROM FormLetters " & _
            "WHERE ID = " & iLetterType
            
    Set oRS = oConn.Execute(cSQL)
    
    
    If Not oRS.EOF Then
        cLetterText = oRS("LetterText")
    Else
        Exit Sub
    End If
        
    
    cSQL = "SELECT ID, Salutation, LastName, FirstName, " & _
            "Address1, Address2, City, State, Zip " & _
            "FROM Requester " & _
            "WHERE Address1 IS NOT NULL " & _
            "ORDER BY LastName"
            
    Set oRS = oConn.Execute(cSQL)

    With VSPrinter1

        .TableBorder = tbNone
        .PageBorder = pbNone
        .Orientation = orPortrait
        .FontName = "Arial"
        .FontSize = 11
        .MarginLeft = 500
        .StartDoc
        
        Do While Not oRS.EOF
        
            cText = cLetterText
        
            cCaseNumber = oRS("id")
            cSalutation = "" & oRS("Salutation")
            cLastName = "" & oRS("lastname")
            cName = oRS("firstname") & " " & oRS("lastname")
            cAddress1 = "" & oRS("Address1")
            cAddress2 = "" & oRS("Address2")
            
            If cAddress2 <> vbNullString Then
                cAddress1 = cAddress1 & vbCrLf & cAddress2
            End If
            
            cCSZ = ("" & oRS("city")) & ", " & oRS("state") & " " & oRS("zip")
            
            cText = Replace(cText, "%Name%", cName)
            cText = Replace(cText, "%Address%", cAddress1)
            cText = Replace(cText, "%CSZ%", cCSZ)
            cText = Replace(cText, "%Salutation%", cSalutation)
            cText = Replace(cText, "%LastName%", cLastName)
            cText = Replace(cText, "%CaseNumber%", cCaseNumber)
            
            .CurrentY = 4320
            
            If iLetterType = ltPlainText Then
                .Text = cText
            Else
                .TextRTF = cText
            End If
            
            oRS.MoveNext
            
            If Not oRS.EOF Then
                .NewPage
            End If
            
        Loop
        
        .EndDoc
        
    End With

    oRS.Close
    Set oRS = Nothing
    
End Sub

Sub Labels()
    Dim oRS As ADODB.Recordset
    Dim cSQL As String
    Dim cName As String
    Dim cAddress1 As String
    Dim cAddress2 As String
    Dim cCSZ As String
    Dim iLine As Integer
    Dim iLabels As Integer
    Dim iCol As Integer
    Dim iCurrentY As Integer
    Dim x As Integer
    
    cSQL = "SELECT LastName, FirstName, Address1, " & _
            "Address2, City, State, Zip " & _
            "FROM Requester " & _
            "WHERE Address1 IS NOT NULL " & _
            "ORDER BY LastName"
            
    Set oRS = oConn.Execute(cSQL)
    
    ' Tracks which column is being printed
    iCol = 1
    
    ' Counter to track the number of labels
    ' printed in a particular column
    iLabels = 0
    
    
    With VSPrinter1
        
        .TableBorder = tbNone
        .PageBorder = pbNone
        .Orientation = orPortrait
        .FontName = "Arial"
        .FontSize = 14
        .MarginTop = 900
        .MarginLeft = 600
        .MarginBottom = 300
        .StartDoc
        
        ' The CurrentY property measures the cursor position in twips
        ' from the top of the page. Because there are 2 columns to
        ' this label style, we will need to return to this
        ' spot when the second columns begin.
        iCurrentY = .CurrentY
        
        Do While Not oRS.EOF
        
            cName = oRS("firstname") & " " & oRS("lastname")
            cAddress1 = "" & oRS("Address1")
            cAddress2 = "" & oRS("Address2")
            cCSZ = ("" & oRS("city")) & ", " & oRS("state") & " " & oRS("zip")
            
            
            'Print an address as a table
            .StartTable
            
            .TableCell(tcCols) = 1
            .TableCell(tcColWidth) = 4000
            
            .TableCell(tcRows) = 7
            .TableCell(tcRowHeight) = 410
            
            
            .TableCell(tcText, 2, 1) = cName
            .TableCell(tcText, 3, 1) = cAddress1
            
            If cAddress2 = vbNullString Then
                .TableCell(tcText, 4, 1) = cCSZ
            Else
                .TableCell(tcText, 4, 1) = cAddress2
                .TableCell(tcText, 5, 1) = cCSZ
            End If
            
            .EndTable
            
                  
            iLabels = iLabels + 1
            
            ' If five labels have been printed, it's time to move
            ' to the top of the page and begin another column
            If iLabels = 5 Then
            
                iLabels = 0
                
                Select Case iCol
                
                    ' If the first has just completed,
                    ' move to the top of the page but go 3900 twips
                    ' to the right
                    Case 1
                        iCol = iCol + 1
                        .CurrentY = iCurrentY
                        .MarginLeft = .MarginLeft + 6000
                    
                    ' If the second column just completed, reset
                    ' the cursor position and eject the page
                    Case 2
                        iCol = 1
                        .MarginLeft = 600
                        .NewPage
                        .CurrentY = iCurrentY
                
                End Select
                
            End If
                        
            oRS.MoveNext
            
        Loop
        
        .EndDoc
        
    End With

    oRS.Close
    Set oRS = Nothing
    
End Sub


Sub StandardColumnarRpt()
    Dim oRS As ADODB.Recordset
    Dim cSQL As String
    Dim x As Long
    
    With VSPrinter1
        
        .HdrFontName = "Arial"
        .HdrFontBold = True
        .HdrFontSize = 12
        
        .Footer = "||Page %d"
        
        .Header = "Seton Software|Sample Report|" & Format(Date, "mm/dd/yyyy")
        
        .FontName = "Arial"
        .FontBold = False
        .FontSize = 10
        
        cSQL = "SELECT LastName, FirstName, State " & _
               "FROM Requester " & _
               "ORDER BY LastName"
        
        Set oRS = oConn.Execute(cSQL)
        
        If oRS.EOF Then
            .KillDoc
            MsgBox "The report criteria you have selected contains no data"
            Exit Sub
        End If
        
        x = 1
        
        .StartDoc
                
        .StartTable
        
        .TableCell(tcCols) = 3
        .TableCell(tcColWidth, , 1) = 2500
        .TableCell(tcColWidth, , 2) = 2000
        .TableCell(tcColWidth, , 3) = 900
        
        Do While Not oRS.EOF
        
            .TableCell(tcInsertRow) = x
            .TableCell(tcText, x, 1) = oRS("LastName")
            .TableCell(tcText, x, 2) = oRS("FirstName")
            .TableCell(tcText, x, 3) = oRS("State")
            
            x = x + 1
            
            oRS.MoveNext
        
        Loop
        
        .EndTable
        
        .EndDoc
        
    End With
    
    oRS.Close
    Set oRS = Nothing
    
End Sub

Sub HeaderFooterRpt()

    Dim x As Integer

    With VSPrinter1
        
        .HdrFontName = "Arial"
        .HdrFontBold = True
        .HdrFontSize = 10
        
        .Footer = "||Page %d"
        
        .Header = "Seton Software Development, Inc.|Sample Report|" & Format(Date, "mm/dd/yyyy")
        
        .StartDoc
        
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
        
    End With
    
End Sub

Sub SampleRpt()
    Dim x As Integer
    
    With VSPrinter1
        
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
        
    End With
    
End Sub


Private Sub VSPrinter1_NewPage()
    Dim bFontBold As Boolean
    
    
    Select Case iReport
    
        Case Report.Sample
            VSPrinter1.AddTable "3900|3900", "Header 1|Header 2", vbNullString
        
        Case 100
            bFontBold = VSPrinter1.FontBold
        
            VSPrinter1.FontBold = True
        
           ' VSPrinter1.AddTable "3900|3900", "Col1|Col2", vbNullString
            'VSPrinter1.AddTable "2500|2000|900", "Last Name|First Name|State", vbNullString
        
            VSPrinter1.FontBold = bFontBold
    
    End Select

End Sub

Sub AddressRpt()

    With VSPrinter1
        
        .StartDoc
        
        .StartTable
        
        .TableCell(tcCols) = 3
        .TableCell(tcRows) = 3
        
        .TableCell(tcColWidth, , 1) = 2000
        .TableCell(tcColWidth, , 2) = 900
        .TableCell(tcColWidth, , 3) = 1800
        
        .TableCell(tcAlign, 1) = taCenterMiddle
        .TableCell(tcRowHeight, 1) = 2500
        .TableCell(tcFont, 1, 1) = "Times New Roman"
        .TableCell(tcFontSize, 1, 1) = 48
        .TableCell(tcFontBold, 1, 1) = True
        .TableCell(tcFontItalic, 1, 1) = True
        .TableCell(tcBackColor, 1, 1) = vbYellow
        .TableCell(tcForeColor, 1, 1) = vbRed
        .TableCell(tcColSpan, 1, 1) = 3
        .TableCell(tcText, 1, 1) = "Joe Smith"
            
        .TableCell(tcFontSize, 2, 1) = 24
        .TableCell(tcColSpan, 2, 1) = 3
        .TableCell(tcText, 2, 1) = "123 Main Street"
        
        .TableCell(tcFontSize, 3) = 24
        .TableCell(tcText, 3, 1) = "Anytown"
        .TableCell(tcText, 3, 2) = "NJ"
        .TableCell(tcText, 3, 3) = "12345"
        
        .EndTable
        
        .EndDoc
        
    End With

End Sub

'Private Sub VSPrinter1_NewPage()
'
'    If VSPrinter1.PageCount Mod 2 = 1 Then
'        VSPrinter1.Orientation = orLandscape
'    Else
'        VSPrinter1.Orientation = orPortrait
'    End If
'
'End Sub


Private Sub VSPrinter1_AfterFooter()
    
    If iReport <> 3 Then
        Exit Sub
    End If

    With VSPrinter1
        .HdrFontName = cHdrFontName
        .HdrFontBold = bHdrFontBold
        .HdrFontSize = iHdrFontSize
    End With
End Sub

Private Sub VSPrinter1_BeforeFooter()

    If iReport <> 3 Then
        Exit Sub
    End If
    
    With VSPrinter1
        cHdrFontName = .HdrFontName
        bHdrFontBold = .HdrFontBold
        iHdrFontSize = .HdrFontSize

        .HdrFontName = "Courier"
        .HdrFontBold = False
        .HdrFontSize = 14
    End With

End Sub

Sub ArrayBasedRpt()
    Dim oRS As ADODB.Recordset
    Dim oExhibitRS As ADODB.Recordset
    Dim oAttendeeRS As ADODB.Recordset
    Dim cSQL As String
    Dim iRow As Integer
    Dim iRowThisPage As Integer
    Dim iMaxPos As Integer
    Dim iTableRow As Integer
    Dim iStart As Integer
    Dim x As Integer
    Dim y As Integer
    Dim lConventionID As Long
    Dim aData(2000, 3) As Variant
    Dim aFormat(2000, 3) As Variant
    
    Const COL_CONVENTION = 1
    Const COL_EXHIBIT = 2
    Const COL_ATTENDEE = 3
    
    cSQL = "SELECT * " & _
           "FROM Convention " & _
           "WHERE ID IN " & _
           "(SELECT ConventionID " & _
           "FROM Attendee)"
    
    Set oRS = oConn.Execute(cSQL)
        
    With VSPrinter1

        .TableBorder = tbNone
        .PageBorder = pbNone
        .Orientation = orLandscape
        .MarginBottom = 500
        .StartDoc
        
        iRowThisPage = 1
        
        Do While Not oRS.EOF
        
            lConventionID = oRS("ID")
            
            'Clear out arrays before beginning new convention
            Erase aData
            Erase aFormat
            
            'After each column is printed, record the maximum
            'value of iRow so the printing logic knows
            'how far down the array to print
            iMaxPos = 0
            
            'Start every convention block at array element 0
            'regardless of where it may begin to print on the page
            iRow = 0
            
            'Convention name should be bold
            aFormat(iRow, COL_CONVENTION) = "b"
            aData(iRow, COL_CONVENTION) = oRS("ConventionName")
            
                
            'Get convention information
            cSQL = "SELECT * " & _
                   "FROM Exhibit " & _
                   "WHERE ConventionID = " & lConventionID
            
            Set oExhibitRS = oConn.Execute(cSQL)
            
            If Not oExhibitRS.EOF Then
                aFormat(iRow, COL_EXHIBIT) = "b"
                aData(iRow, COL_EXHIBIT) = "EXHIBITS"
                iRow = iRow + 1
            End If
            
            Do While Not oExhibitRS.EOF
            
                aData(iRow, COL_EXHIBIT) = oExhibitRS("Descr")
            
                iRow = iRow + 1
            
                oExhibitRS.MoveNext
            
            Loop
            
            If iMaxPos < iRow Then
                iMaxPos = iRow
            End If
        
        
        
            'Now get attendee information
            iRow = 0
        
            cSQL = "SELECT * " & _
                   "FROM Attendee " & _
                   "WHERE ConventionID = " & lConventionID
            
            Set oAttendeeRS = oConn.Execute(cSQL)
            
            If Not oAttendeeRS.EOF Then
                aFormat(iRow, COL_ATTENDEE) = "b"
                aData(iRow, COL_ATTENDEE) = "ATTENDEES"
                iRow = iRow + 1
            End If
            
            Do While Not oAttendeeRS.EOF
            
                aData(iRow, COL_ATTENDEE) = oAttendeeRS("name")
            
                iRow = iRow + 1
            
                oAttendeeRS.MoveNext
            
            Loop
            
            If iMaxPos < iRow Then
                iMaxPos = iRow
            End If
            
            oRS.MoveNext
            


            If .CurrentY > 10000 Then
            
                iRowThisPage = 1
    
                Call DrawLines(iStart, .CurrentY)
                
                .NewPage
                
            End If
                
            .TableBorder = tbNone
            
            .StartTable
            .TableCell(tcCols) = 3
                  
            'Count what row we're on for the convention currently being
            'printed. This is the row counter for the StartTable/EndTable
            'set and is reset for each convention regardless
            'of how many conventions print on a single page
            iTableRow = 1
                             
            
            iStart = .CurrentY
                
            For x = 0 To iMaxPos

                    
                'Avoid printing a blank line at the top of the next
                'page just before beginning a new convention
                If IsEmpty(aData(iMaxPos, COL_CONVENTION)) And _
                   IsEmpty(aData(iMaxPos, COL_EXHIBIT)) And _
                   IsEmpty(aData(iMaxPos, COL_ATTENDEE)) And _
                   iRowThisPage = 1 And _
                   x = iMaxPos Then
                   
                    Exit For
                
                End If
                    
                
                .TableCell(tcInsertRow) = iTableRow
                
                .TableCell(tcFontSize, iTableRow) = 10
                
                .TableCell(tcColWidth, iTableRow, COL_CONVENTION) = 3500
                .TableCell(tcColWidth, iTableRow, COL_EXHIBIT) = 3500
                .TableCell(tcColWidth, iTableRow, COL_ATTENDEE) = 3500
                
                For y = COL_CONVENTION To COL_ATTENDEE
                
                    If InStr(aFormat(x, y), "b") <> 0 Then
                        .TableCell(tcFontBold, iTableRow, y) = True
                    End If
                    
                    .TableCell(tcText, iTableRow, y) = "" & aData(x, y)
                    
                Next y
        
                iTableRow = iTableRow + 1
                iRowThisPage = iRowThisPage + 1
        
                'No more than 36 rows per page
                If iRowThisPage > 36 Then
                    .EndTable
                    
                    iTableRow = 1
                    iRowThisPage = 1
        
                    Call DrawLines(iStart, .CurrentY)
                    
                    .NewPage
                    
                    iStart = .CurrentY
        
                    .StartTable
                    .TableCell(tcCols) = 3
                
                End If
                
            Next x
    
            .EndTable
                
            Call DrawLines(iStart, .CurrentY)
        
        Loop
        
        .EndDoc
        
        oRS.Close
        Set oRS = Nothing
    
    End With
    
End Sub
            
Sub DrawLines(iStart As Integer, iEnd As Integer)

    With VSPrinter1
        .PenWidth = 15
        .BrushStyle = bsTransparent
        .DrawRectangle .MarginLeft, iStart, 10505, iEnd
        .DrawRectangle 3500 + .MarginLeft, iStart, 3505 + .MarginLeft, iEnd
        .DrawRectangle 7000 + .MarginLeft, iStart, 7005 + .MarginLeft, iEnd
    End With
    
End Sub

Sub SignatureDisplay()

    With VSPrinter1
        .StartDoc
        .DrawPicture Picture1, 6000, 4000, "50%", "50%"
        .EndDoc
    End With
    
End Sub

Sub CrossTabRpt()
    Dim oRS As ADODB.Recordset
    Dim cSQL As String
    Dim x As Integer
    Dim y As Integer
    Dim Z As Integer
    Dim iSize As Integer
    Dim iColumnBegin As Integer
    Dim iProdAtPageStart  As Integer
    Dim iProdRow As Integer
    Dim iPageRow As Integer
    Dim iPageDataSet As Integer
    Dim iPagesPerRow As Integer
    Dim iPagesPerDataSet As Integer
    Dim iProdSet As Integer
    Dim iRow As Integer
    Dim iMaxCol As Integer
    Dim iMaxRow As Integer
    Dim iProdCnt As Integer
    Dim iSourceCnt As Integer
    Dim lGrandTotal As Long
    Dim lTotal As Long
    Dim iColPos As Integer
    Dim iProdPos As Integer
    Dim iSourcePos As Integer
    Dim iSourceSize As Integer
    Dim lSourceID As Long
    Dim aProduct() As Variant
    Dim aSource() As Variant
    Dim aData() As Long
        
    'Should be defined so as to allow room for a totals
    'column to the right of the last column in a row
    Const COLS_PER_PAGE = 5
    
    'Should be defined so as to allow room for a totals
    'column below the last row
    Const ROWS_PER_PAGE = 7
                            
                            
    'Count the number of products
    
    cSQL = "SELECT COUNT(*) " & _
           "FROM product "
           
    Set oRS = oConn.Execute(cSQL)
    
    iProdCnt = oRS.Fields(0)
    
    
    'Count the number of sources
    
    cSQL = "SELECT COUNT(*) " & _
           "FROM source"
           
    Set oRS = oConn.Execute(cSQL)
    
    iSourceCnt = oRS.Fields(0)
    
    
    'Resize array to hold report data
    ReDim aProduct(iProdCnt, 1)
    ReDim aSource(iSourceCnt, 1)
    ReDim aData(iProdCnt, iSourceCnt)
    
    
    
    'Load product ids and names into arrays
    
    cSQL = "SELECT id, descr " & _
       "FROM product " & _
       "ORDER BY descr"
       
    Set oRS = oConn.Execute(cSQL)
    
    x = 1
    
    Do While Not oRS.EOF
    
        aData(x, 0) = oRS("id")
        aProduct(x, 0) = oRS("id")
        aProduct(x, 1) = oRS("descr")
        
        x = x + 1
        
        oRS.MoveNext
    
    Loop
    
    
    'Load source ids and names into arrays
    
    cSQL = "SELECT id, descr " & _
       "FROM source " & _
       "ORDER BY descr"
       
    Set oRS = oConn.Execute(cSQL)
    
    x = 1
    
    Do While Not oRS.EOF
        
        aData(0, x) = oRS("id")
        aSource(x, 0) = oRS("id")
        aSource(x, 1) = oRS("descr")
        
        x = x + 1
        
        oRS.MoveNext
        
    Loop
    
    
    
    'Use GROUP BY query to extract summary data by product and source
    iSourceSize = UBound(aData, 2)
    
    cSQL = "SELECT productid, sourceid, COUNT(sourceid) AS SourceCnt " & _
           "FROM requester " & _
           "WHERE sourceid IS NOT NULL " & _
           "GROUP BY productid, sourceid"
    
    Set oRS = oConn.Execute(cSQL)
    
    
    'Load this data into the appropriate array coordinates so we end up with
    'an array that matches the layout of the final report
    Do While Not oRS.EOF
        
        'Find the product vertically
        iProdPos = AScan(aData, oRS("productid"), 0)
        
        
        'Find the source horizontally
        For iSourcePos = 1 To iSourceSize
        
            If aData(0, iSourcePos) = oRS("sourceid") Then
                Exit For
            End If
        
        Next iSourcePos
        
        'Place the value if the assigned location
        aData(iProdPos, iSourcePos) = oRS("SourceCnt")
        
        oRS.MoveNext
        
    Loop
    

    
    With VSPrinter1
        .TableBorder = tbNone
        .PageBorder = pbNone
        .Orientation = orLandscape
        .FontName = "Arial"
        .FontSize = 10
        .StartDoc
    End With
    
    
    'How many pages are needed to print an entire row of data
    iPagesPerRow = Int(iSourceCnt / COLS_PER_PAGE) + 1
    
    
    'How many pages are needed to print a full set of products A through Z
    iPagesPerDataSet = Int(iProdCnt / ROWS_PER_PAGE)
    
    If iProdCnt Mod ROWS_PER_PAGE <> 0 Then
        iPagesPerDataSet = Int(iPagesPerDataSet + 1)
    End If
    
    
    'Which product array element are we on at the beginning of a given page
    iProdAtPageStart = 1
    
    'Of the number of pages it takes to print all the products,
    'as indicated by iPagesPerDataSet, which pages are we currently on
    iProdSet = 1
                 
    
    With VSPrinter1
        
        'Row counter for the table on the page currently being
        'printed. Reset after each page.
        iRow = 1
        
        'Page counter for rows which span multiple pages
        iPageRow = 1
        
        'Page counter for data sets which span multiple pages
        iPageDataSet = 1
        
        'Position in the array where the data should begin printing. The product
        'description begins in column 1 on each page. Though its a zero-based
        'array we are ignoring column zero
        iColumnBegin = 2
        
        'Which product of the entire list of products are we on
        iProdRow = 1
        
        'Print the header across the top of the page
        Call CrossTabHdr(aSource(), iColumnBegin, iPageRow, COLS_PER_PAGE, iSourceCnt)
                
                
        .StartTable
                
        'Always one more than the COLS_PER_PAGE so as to display the product names
        'down the left side of the page.
        .TableCell(tcCols) = COLS_PER_PAGE + 2
        
        Do While iProdRow <= iProdCnt
    
            'Insert row and fill with product name
            .TableCell(tcInsertRow) = iRow
            .TableCell(tcColWidth, iRow, 1) = "3400"
            .TableCell(tcText, iRow, 1) = aProduct(iProdRow, 1)
            
            
            'Calculate how many columns are being printed across this page
            iMaxCol = (COLS_PER_PAGE * iPageRow) + 1
            
            If iMaxCol > iSourceCnt Then
                iMaxCol = iSourceCnt + 1
            End If
            
            If iPageRow = 1 Then
                iColumnBegin = 2
            Else
                iColumnBegin = COLS_PER_PAGE * (iPageRow - 1) + 2
            End If
            
            'Now print those columns - one row's worth of data
            iColPos = 1
            
            For Z = iColumnBegin To iMaxCol
    
                iColPos = iColPos + 1
    
                .TableCell(tcColWidth, iRow, iColPos) = "1600"
                .TableCell(tcAlign, iRow, iColPos) = taRightMiddle
                .TableCell(tcText, iRow, iColPos) = aData(iProdRow, Z - 1)
                
                lGrandTotal = lGrandTotal + aData(iProdRow, Z - 1)
    
            Next Z
            
            'Go to the next row for this page
            iRow = iRow + 1
            
            'Go to the next product in this data set
            iProdRow = iProdRow + 1
            
            'If we have more rows than are allowed on a page or
            'we've run out of products because we're on the last page
            If iRow > ROWS_PER_PAGE Or iProdRow > iProdCnt Then
                        
                'If its the last part of the data set, then
                'print column totals across bottom of page
                If iPageDataSet = iPagesPerDataSet Then
                
                    .TableCell(tcInsertRow) = iRow
    
                    .TableCell(tcColWidth, iRow, 1) = "3400"
                    .TableCell(tcText, iRow, 1) = "Totals"
                    .TableCell(tcFontBold, iRow, 1) = True
                    .TableCell(tcAlign, iRow, 1) = taLeftMiddle
    
                    iColPos = 1
                    
                    If iPageRow = 1 Then
                        iColumnBegin = 2
                    Else
                        iColumnBegin = COLS_PER_PAGE * (iPageRow - 1) + 2
                    End If
            
                    For Z = iColumnBegin To iMaxCol
    
                        iColPos = iColPos + 1
                        
                        .TableCell(tcColWidth, iRow, iColPos) = "1600"
                        .TableCell(tcAlign, iRow, iColPos) = taRightMiddle
                        .TableCell(tcText, iRow, iColPos) = ASum(aData, 1, , Z - 1)
    
                    Next Z
                
                End If
                            
                
                'If we've printed one full row's worth of data for a data set
                If iPageRow = iPagesPerRow Then
                
                    'Print total column for the row if this is the
                    'last page on which a row prints
                    If iSourceCnt Mod COLS_PER_PAGE = 0 Then
                        iColPos = 2
                    Else
                        iColPos = (iSourceCnt Mod COLS_PER_PAGE) + 2
                    End If
                    
                    iMaxRow = (iProdAtPageStart + ROWS_PER_PAGE) - 1
                    
                    If iMaxRow > iProdCnt Then
                        iMaxRow = iProdCnt
                    End If
                    
                    iRow = 1
                
                    For Z = iProdAtPageStart To iMaxRow
    
                        .TableCell(tcColWidth, iRow, iColPos) = "1600"
                        .TableCell(tcAlign, iRow, iColPos) = taRightMiddle
                        .TableCell(tcText, iRow, iColPos) = ASum(aData, 1, iSourceCnt, , Z)
                        
                        iRow = iRow + 1
    
                    Next Z
                    
                    'If this is the very last page print the grand total
                    If iProdRow >= iProdCnt And _
                       iPageDataSet = iPagesPerDataSet Then
    
                        .TableCell(tcColWidth, iRow, iColPos) = "1600"
                        .TableCell(tcAlign, iRow, iColPos) = taRightMiddle
                        .TableCell(tcText, iRow, iColPos) = lGrandTotal
    
                    End If
                                            
                
                    .EndTable
                    
                    'If this is the very last page they exit the
                    'loop and close the Recordset
                    If iProdRow >= iProdCnt And _
                       iPageDataSet = iPagesPerDataSet Then
                        Exit Do
                    End If
                    
                    .NewPage
                
                    iPageDataSet = iPageDataSet + 1
                    iProdRow = iProdSet * ROWS_PER_PAGE
                    iProdSet = iProdSet + 1
                    iPageRow = 1
                    iProdAtPageStart = iProdRow
                    iColumnBegin = 2
        
                Else
                
                    .EndTable
                    
                    .NewPage
                
                    iPageRow = iPageRow + 1
                    iProdRow = iProdAtPageStart
                    iColumnBegin = (COLS_PER_PAGE * (iPageRow - 1)) + 2
                
                End If
                
                iRow = 1
                    
                'Print the header
                Call CrossTabHdr(aSource(), iColumnBegin, iPageRow, COLS_PER_PAGE, iSourceCnt)
                
                .StartTable
                
                iMaxCol = iSourceCnt - (COLS_PER_PAGE * (iPageRow - 1))
                
                If iMaxCol > COLS_PER_PAGE Then
                    iMaxCol = COLS_PER_PAGE
                End If
            
                .TableCell(tcCols) = COLS_PER_PAGE + 2
                
            End If
        
        Loop
        
        .EndDoc
        
    End With

    oRS.Close
    Set oRS = Nothing

End Sub

Sub CrossTabHdr(aSource() As Variant, iColumnBegin As Integer, iPageRow As Integer, iColsPerPage As Integer, iSourceCnt As Integer)
    Dim x As Integer
    Dim iColPos As Integer
    Dim iMaxCol As Integer
    Dim cFmt As String
    Dim cTitle As String
    
    cFmt = "3400"
    cTitle = "Product|"

    iMaxCol = (iColsPerPage * iPageRow)
    
    If iMaxCol > iSourceCnt Then
        iMaxCol = iSourceCnt
    End If
            
    For x = iColumnBegin - 1 To iMaxCol
    
        iColPos = x - (iColsPerPage * (iPageRow - 1))

        cFmt = cFmt & "|>1600"

        cTitle = cTitle & aSource(x, 1) & "|"
    
    Next x
    
    If iMaxCol = iSourceCnt Then

        cFmt = cFmt & "|>1600"

        cTitle = cTitle & "Totals"
        
    End If
        
    With VSPrinter1

        .FontBold = True
        .AddTable cFmt, "", cTitle
        .FontBold = False
        
        .TableBorder = tbNone
        
    End With

End Sub


Function AScan(aArray, uValue, iCol) As Integer
    ' This function scans a passed Array to
    ' determine if a passed value (uValue)
    ' exists in the passed Array (aArray)
    ' If it does it will return the index pointer
    ' Otherwise it will return -1
    
    Dim iOffset As Integer  ' Array offset pointer
    Dim iReturn As Integer  ' Return value
    Dim nSize   As Integer  ' Avoid multiple calls to ALen()
      
      ' Set return value
    iReturn = -1
      
      ' Store the Array Size
    nSize = UBound(aArray)
    
       ' Greater than 0 so process all elements
    For iOffset = 0 To nSize
        If aArray(iOffset, iCol) = uValue Then
            iReturn = iOffset
            Exit For
        End If
    Next iOffset
      
    AScan = iReturn
        
End Function

Function ASum(aArray As Variant, Optional vntStart As Variant, Optional vntEnd As Variant, Optional vntCol As Variant, Optional vntRow As Variant) As Double
    Dim x As Integer
    Dim dblResult As Double
    
    If IsMissing(vntStart) Then
        vntStart = 0
    End If
    
    If IsMissing(vntEnd) Then
    
        If Not IsMissing(vntRow) Then
            vntEnd = UBound(aArray, 2)
        Else
            vntEnd = UBound(aArray, 1)
        End If
    
    End If
    
    If Not IsMissing(vntRow) Then
    
        For x = vntStart To vntEnd
            dblResult = dblResult + CDbl(IIf(aArray(vntRow, x) = vbNullString, 0, aArray(vntRow, x)))
        Next
    
    Else
    
        For x = vntStart To vntEnd
            dblResult = dblResult + CDbl(IIf(aArray(x, vntCol) = vbNullString, 0, aArray(x, vntCol)))
        Next
    
    End If
    
    ASum = dblResult
    
End Function

Sub CrossTabSQL()
    Dim oConn As New ADODB.Connection
    Dim oRS As ADODB.Recordset
    Dim cConnectString As String
    Dim cColumn As String
    Dim cSQL As String

    cConnectString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;" & _
                     "Persist Security Info=False;" & _
                     "Initial Catalog=EnterpriseReports;" & _
                     "Data Source=(local)"
    
    With oConn
        .ConnectionString = cConnectString
        .Open
    End With
    
    cSQL = "SELECT * FROM Source"
    
    Set oRS = oConn.Execute(cSQL)
    
    
    cSQL = "SELECT p.Descr,"
    
    Do While Not oRS.EOF
    
        cColumn = oRS("descr")
        
        cColumn = Replace(cColumn, Space(1), vbNullString)
        cColumn = Replace(cColumn, "-", vbNullString)
    
        cSQL = cSQL & "SUM(CASE SourceID WHEN " & oRS("id") & _
                      " THEN 1 ELSE 0 END) AS " & cColumn & ","
    
        oRS.MoveNext
    
    Loop
    
    cSQL = Mid$(cSQL, 1, Len(cSQL) - 1)
    
    cSQL = cSQL & " FROM Requester r, Product p " & _
                  "WHERE r.ProductID = p.ID " & _
                  "GROUP BY p.Descr"
                  
    MsgBox cSQL
    
End Sub


Sub ExecSQL()
    Dim oConn As New ADODB.Connection
    Dim oCmd As New ADODB.Command
    Dim oRS As ADODB.Recordset
    Dim cConnectString As String
    Dim cSQL As String

   ' cConnectString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;" & _
                     "Persist Security Info=False;" & _
                     "Initial Catalog=EnterpriseReports;" & _
                     "Data Source=Y8P8P"
    
    cConnectString = "Provider=SQLOLEDB.1;Password=Ovaltine" & _
                            ";Persist Security Info=True;User ID=sa" & _
                            ";Initial Catalog=EnterpriseReports;" & _
                            ";Data Source=Y8P8P"
                            
    With oConn
        .ConnectionString = cConnectString
        .Open
    End With
    
    cSQL = "SELECT * FROM Source"
    
    With oCmd
        Set .ActiveConnection = oConn
        .CommandText = "sp_ExecSQL"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 20
        .Parameters.Append .CreateParameter("SQL", adVarChar, adParamInput, 2000, cSQL)
        Set oRS = .Execute
    End With
        
End Sub

Sub ExportReportToPDF()

    With VSPDF81
        .Title = "PDF Export"
        .Creator = "Your Name"
        .Author = "My Name"
        .Subject = "Enterprise Reports with VB6 and VB.NET"
        .Keywords = "VB VB.NET Reports"
        .ConvertDocument VSPrinter1, App.Path & "\myreport.pdf"
    End With

End Sub

Sub JumpHyperlinks()
    Dim x As Integer

    With VSPrinter1
    
        .AutoLinkNavigate = True
    
        .StartDoc
        
        .Preview = True
        
        .StartTable
                
        .TableCell(tcCols) = 2
        .TableCell(tcColWidth, , 1) = 3900
        .TableCell(tcColWidth, , 2) = 3900
                
        Call .AddLinkTarget("US Presidents", "TopOfFirstPage")
        
        .Text = "" & vbCrLf
        .Text = "" & vbCrLf
        
        Call .AddLink("George Washington", "www.microsoft.com", True)
    
        .Text = "" & vbCrLf
        
        For x = 1 To 75
        
            .Text = "More text" & vbCrLf
        
        Next x
        
        .EndTable
        
        
        Call .AddLink("Go to the beginning of the document", "#TopOfFirstPage", True)
        
        .EndDoc
        
    End With

    With VSPDF81
        .Title = "PDF Export"
        .Creator = "Your Name"
        .Author = "My Name"
        .Subject = "Enterprise Reports with VB6 and VB.NET"
        .Keywords = "VB VB.NET Reports"
        .ConvertDocument VSPrinter1, App.Path & "\myreport.pdf"
    End With
        
End Sub

Private Sub VSPrinter1_MouseLink(Link As String, ByVal Clicked As Boolean, Cancel As Integer)
    
    If Not Clicked Then
        VSPrinter1.ToolTipText = Link
    Else
        VSPrinter1.ToolTipText = vbNullString
    End If

End Sub
