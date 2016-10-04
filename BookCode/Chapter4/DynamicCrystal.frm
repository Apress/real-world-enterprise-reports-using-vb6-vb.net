VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   9615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11415
      lastProp        =   500
      _cx             =   20135
      _cy             =   16960
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   9840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            
Private Sub Form_Load()
    Dim objApplication As New CRAXDRT.Application
    Dim objReport As New CRAXDRT.Report
    Dim objText As CRAXDRT.TextObject
    Dim objField As CRAXDRT.FieldObject
    Dim objFieldDef As CRAXDRT.DatabaseFieldDefinition
    Dim objSection As CRAXDRT.Section
    Dim oConn As ADODB.Connection
    Dim oCmd As ADODB.Command
    Dim cConnectString As String
    Dim cSQL As String
    Dim cCaption As String
    Dim cFieldName As String
    Dim iFieldCnt As Integer
    Dim iLeft As Integer
    Dim x As Integer
    
    'Create ADO Connection object
    Set oConn = New ADODB.Connection
    cConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                     "Data Source=C:\BookCode\SampleDatabase.mdb;" & _
                     "Persist Security Info=False"
    oConn.Open cConnectString
    
    
    cSQL = "SELECT LastName, FirstName, Address1, City, State " & _
           "FROM Requester " & _
           "WHERE State IN ('NJ', 'NY', 'SC')"
    
    
    'Create ADO Command object
    Set oCmd = New ADODB.Command
    Set oCmd.ActiveConnection = oConn
    oCmd.CommandText = cSQL
    oCmd.CommandType = adCmdText
    
    'Instantiate new Crystal Application object
    Set objApplication = New CRAXDRT.Application
    
    'Instantiate a blank report object
    Set objReport = objApplication.NewReport
         
    'Set the data source of the new report
    Call objReport.Database.AddADOCommand(oConn, oCmd)
    objReport.PaperSize = crPaperLetter
    objReport.TopMargin = 200
    objReport.BottomMargin = 200
    
    
    cCaption = "US Presidents - Where are they now?"
    
    'Create report title section
    objReport.Sections(1).Height = objReport.Sections(1).Height + 300
    Set objText = objReport.Sections(1).AddTextObject(cCaption, 1700, 0)
    
    With objText
        .HorAlignment = crHorCenterAlign
        .BorderColor = vbBlack
        .Font.Size = 18
        .Font.Bold = True
        .TextColor = vbBlue
        .Height = 600
        .Width = 7500
    End With
    
    
    With objReport
    
        'How many fields in data source
        iFieldCnt = .Database.Tables(1).Fields.Count
            
            
        'Loop through each field and add it to the report
        
        For x = 1 To iFieldCnt
        
            cFieldName = .Database.Tables(1).Fields(x).Name
        
            Set objField = .Sections(3).AddFieldObject(cFieldName, iLeft, 0)
            
            objField.Font.Name = "Ariel"
            objField.Font.Size = 10
                        
            iLeft = iLeft + (.Database.Tables(1).Fields(x).NumberOfBytes * 60)
        
        Next x
          
        'Group by state field
        Set objFieldDef = .Database.Tables(1).Fields(5)
                
        Call objReport.AddGroup(0, objFieldDef, crGCAnyValue, crAscendingOrder)
        
        Set objSection = objReport.Sections.Item("GH1")
        Call objSection.AddFieldObject(objFieldDef, 100, 0)
        
        
    
        'Add a subtotal count
        Set objSection = objReport.Sections.Item("RF")
                    
        objSection.BackColor = vbCyan
        
        cCaption = "Total US Presidents listed:"
        
        Set objText = objSection.AddTextObject(cCaption, 100, 0)
        
        objText.Font.Name = "Ariel"
        objText.Font.Size = 10
        objText.Font.Bold = True
        objText.Width = 3000
        
        Set objFieldDef = .Database.Tables(1).Fields(1)
        
        Call objSection.AddSummaryFieldObject(objFieldDef, crSTCount, 2500, 0)
    
    End With
    
    'Display the report
    With CRViewer91
        .ReportSource = objReport
        .Zoom (100)
        .ViewReport
    End With
    
    'KB - c2009297
    
'    With objReport.ExportOptions
'        .DiskFileName = App.Path & "\myreport.pdf"
'        .DestinationType = crEDTDiskFile
'        .PDFExportAllPages = True
'        .FormatType = crEFTPortableDocFormat
'    End With
'
'    objReport.Export False
'
'
'    With objReport.ExportOptions
'        .XMLFileName = App.Path & "\myreport.xml"
'        .DestinationType = crEDTDiskFile
'        .FormatType = crEFTXML
'    End With
'
'    objReport.Export False
'
'
'    With objReport.ExportOptions
'        .DiskFileName = App.Path & "\myreport.xls"
'        .DestinationType = crEDTDiskFile
'        .FormatType = crEFTExcel80
'    End With
'
'    objReport.Export False
'
'
'    With objReport.ExportOptions
'        .DiskFileName = App.Path & "\myreport.doc"
'        .DestinationType = crEDTDiskFile
'        .FormatType = crEFTWordForWindows
'    End With
'
'    objReport.Export False
'
'
'    With objReport.ExportOptions
'        .HTMLFileName = App.Path & "\myreport.html"
'        .DestinationType = crEDTDiskFile
'        .FormatType = crEFTHTML40
'    End With
'
'    objReport.Export False
            
End Sub
