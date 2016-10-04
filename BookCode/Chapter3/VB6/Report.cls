VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim frmReportViewer As Form
Dim cReport As String
Dim cReportTitle As String
Dim cReportDateFrom As String
Dim cReportDateTo As String
Dim cReportSQL As String
Dim iHelpContext As Integer
Dim iPrintDestination As Integer
Dim bPrintCriteriaPage As Boolean
Dim cReportCriteria As String
Dim cLocalPath As String

Public Property Set ReportViewer(ByVal vData As Object)
    Set frmReportViewer = vData
End Property

Public Property Let Report(ByVal vData As String)
    cReport = vData
End Property

Public Property Get Report() As String
    Report = cReport
End Property

Public Property Let HelpContext(ByVal vData As Integer)
    iHelpContext = vData
End Property

Public Property Get HelpContext() As Integer
    HelpContext = iHelpContext
End Property

Public Property Let PrintDestination(ByVal vData As Integer)
    iPrintDestination = vData
End Property

Public Property Get PrintDestination() As Integer
    PrintDestination = iPrintDestination
End Property

Public Property Get ReportTitle() As String
    ReportTitle = cReportTitle
End Property

Public Property Let ReportTitle(ByVal vData As String)
    cReportTitle = vData
End Property

Public Property Get ReportDateFrom() As String
    ReportDateFrom = cReportDateFrom
End Property

Public Property Let ReportDateFrom(ByVal vData As String)
    cReportDateFrom = vData
End Property

Public Property Get ReportDateTo() As String
    ReportDateTo = cReportDateTo
End Property

Public Property Let ReportDateTo(ByVal vData As String)
    cReportDateTo = vData
End Property

Public Property Get ReportSQL() As String
    ReportSQL = cReportSQL
End Property

Public Property Let ReportSQL(ByVal vData As String)
    cReportSQL = vData
End Property

Public Property Get ReportCriteria() As String
    ReportCriteria = cReportCriteria
End Property

Public Property Let ReportCriteria(ByVal vData As String)
    cReportCriteria = vData
End Property

Public Property Get PrintCriteriaPage() As Boolean
    PrintCriteriaPage = bPrintCriteriaPage
End Property

Public Property Let PrintCriteriaPage(ByVal vData As Boolean)
    bPrintCriteriaPage = vData
End Property

Public Property Get LocalPath() As String
    LocalPath = cLocalPath
End Property

Public Property Let LocalPath(ByVal vData As String)
    cLocalPath = vData
End Property

Public Sub ShowPrint()

    With frmPrint
        Set .objReport = Me
        .Show vbModal
    End With
    
End Sub

Sub DurationByStaffRpt()

    With frmReportViewer!VSPrinter1
        .TableBorder = tbNone
        .PageBorder = pbNone
        .Orientation = orPortrait
        .FontName = "Arial"
        .StartDoc
        
        .Text = "Hi, I'm the " & cReport & " method of the Report class" & vbCrLf & vbCrLf
        
        .Text = "The criteria for this report is " & vbCrLf & cReportCriteria
        
        .EndDoc
        
    End With
    
End Sub

Sub DurationByProductRpt()

    With frmReportViewer!VSPrinter1
        .TableBorder = tbNone
        .PageBorder = pbNone
        .Orientation = orPortrait
        .FontName = "Arial"
        .StartDoc
        
        .Text = "Hi, I'm the " & cReport & " method of the Report class" & vbCrLf & vbCrLf
        
        .Text = "The criteria for this report is " & vbCrLf & cReportCriteria
        
        .EndDoc
        
    End With
    
End Sub

Sub DurationBySourceRpt()
    Dim cSQL As String
    
    With frmReportViewer!VSPrinter1
        .TableBorder = tbNone
        .PageBorder = pbNone
        .Orientation = orPortrait
        .FontName = "Arial"
        .StartDoc
        
        cSQL = "SELECT * " & _
               "FROM inquiry " & _
               IIf(cReportSQL = vbNullString, vbNullString, cReportSQL)
        
        .Text = "Hi, I'm the " & cReport & " method of the Report class" & vbCrLf & vbCrLf
        
        .Text = "The criteria for this report is " & vbCrLf & cReportCriteria
        
        .EndDoc
        
    End With
    
End Sub
