VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Office XP"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   6435
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRunDemo 
      Caption         =   "Excel Pivot Table"
      Height          =   495
      Index           =   10
      Left            =   2280
      TabIndex        =   11
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdRunDemo 
      Caption         =   "Word Styles"
      Height          =   495
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdRunDemo 
      Caption         =   "Word Objects"
      Height          =   495
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton cmdRunDemo 
      Caption         =   "Word Overview"
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdRunDemo 
      Caption         =   "Lotus Notes Email"
      Height          =   495
      Index           =   6
      Left            =   4440
      TabIndex        =   6
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdRunDemo 
      Caption         =   "Outlook Email"
      Height          =   495
      Index           =   5
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdRunDemo 
      Caption         =   "Excel Graph"
      Height          =   495
      Index           =   4
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdRunDemo 
      Caption         =   "Excel Report"
      Height          =   495
      Index           =   3
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdRunDemo 
      Caption         =   "Word Mailing Labels"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdRunDemo 
      Caption         =   "Word Mail Merge"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdRunDemo 
      Caption         =   "Word Tables"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oConn As New ADODB.Connection

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRunDemo_Click(Index As Integer)
    
    Select Case Index
    
        Case 0
            Call WordTables
        
        Case 1
            Call MailMerge
            
        Case 2
            Call MailingLabels
            
        Case 3
            Call ExcelReport
            
        Case 4
            Call ExcelGraph
            
        Case 5
            Call OutlookEmail
            
        Case 6
            Call LotusNotesEmail("Seton.software@verizon.net", "joesmith@xyz.com", "My Subject", "Important message")
            
        Case 7
            Call WordOverview
        
        Case 8
            Call WordObjects
            
        Case 9
            Call WordStyles
            
        Case 10
            Call PivotTable
            
    End Select
    
End Sub

Sub WordOverview()
    Dim objWord As Word.Application
    Dim objDocument As Word.Document
    Dim cStats As String
    
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False
    
    Set objDocument = objWord.Documents.Open(App.Path & "\WordExample.doc")
    
    With objDocument
    
        cStats = "# paragraphs: " & .Paragraphs.Count & vbCrLf
        cStats = cStats & "# sentences: " & .Sentences.Count & vbCrLf
        cStats = cStats & "# words: " & .Words.Count & vbCrLf
        cStats = cStats & "# characters: " & .Characters.Count & vbCrLf & vbCrLf
        
        cStats = cStats & "First Paragraph: " & .Paragraphs(1).Range.Text & vbCrLf
        cStats = cStats & "Second Sentence: " & .Sentences(2) & vbCrLf
        cStats = cStats & "Third Word: " & .Words(3) & vbCrLf
        cStats = cStats & "Fourth Character: " & .Characters(4) & vbCrLf
        
    End With
        
    MsgBox cStats

End Sub

Sub WordObjects()
    Dim objWord As Word.Application
    Dim objDocument As Word.Document
    Dim objRange As Word.Range
    
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False
    
    Set objDocument = objWord.Documents.Open(App.Path & "\WordExample.doc")
    
    With objDocument
    
        Call objDocument.Sentences(3).Copy
        
        Set objRange = .Paragraphs(2).Range
        
    End With

End Sub

Sub WordStyles()
    Dim objWord As Word.Application
    Dim objDocument As Word.Document
    Dim objRange As Word.Range
    Dim objStyle As Word.Style
    
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True
    
    Set objDocument = objWord.Documents.Open(App.Path & "\WordExample.doc")
    
    Set objStyle = objDocument.Styles.Add("Chapter Title 1")
    
    With objStyle
        .Font.Name = "Arial"
        .Font.Size = 28
        .Font.Bold = True
        .Font.Color = wdColorDarkRed
    End With
    
    Set objRange = objDocument.Paragraphs(1).Range
    
    objRange.Style = "Chapter Title 1"
    
End Sub

Sub WordTables()
    Dim oRS As ADODB.Recordset
    Dim objWord As Word.Application
    Dim objDocument As Word.Document
    Dim objTable As Word.Table
    Dim objRange As Word.Range
    Dim X As Integer
    Dim iFieldCnt As Integer
    Dim iRow As Integer
    Dim cSQL As String
    
    cSQL = "SELECT LastName, FirstName, State " & _
           "FROM Requester " & _
           "ORDER BY LastName"
    
    Set oRS = oConn.Execute(cSQL)
    
    iFieldCnt = oRS.Fields.Count
    iRow = 1
    
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True
    
    Set objDocument = objWord.Documents.Add
    
    With objDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range
        .InsertAfter "Status Report"
        .Paragraphs.Alignment = wdAlignParagraphLeft
    End With
    
    objDocument.Sections(1).Footers(wdHeaderFooterPrimary). _
        PageNumbers.Add.Alignment = wdAlignPageNumberCenter
    
    With objDocument.PageSetup
        .LeftMargin = InchesToPoints(0.5)
        .RightMargin = InchesToPoints(0.5)
    End With
    
    Set objRange = objWord.ActiveDocument.Range(0, 0)
    
    Set objTable = objDocument.Tables.Add(objRange, 1, iFieldCnt)
        
    With objTable
        .Cell(iRow, 1).Range.InsertAfter "Last Name"
        .Cell(iRow, 1).Width = InchesToPoints(1.2)
        
        .Cell(iRow, 2).Range.InsertAfter "First Name"
        .Cell(iRow, 2).Width = InchesToPoints(1.2)
                        
        .Cell(iRow, 3).Range.InsertAfter "State"
        .Cell(iRow, 3).Width = InchesToPoints(1.5)
        .Cell(iRow, 3).LeftPadding = PixelsToPoints(30)
    
        Call .Rows.Add
    
        iRow = iRow + 1
            
        Do While Not oRS.EOF
                            
            .Cell(iRow, 1).Range.InsertAfter oRS("LastName")
            .Cell(iRow, 1).Width = InchesToPoints(1.2)
            
            .Cell(iRow, 2).Range.InsertAfter oRS("FirstName")
            .Cell(iRow, 2).Width = InchesToPoints(1.2)
                        
            .Cell(iRow, 3).Range.InsertAfter oRS("State")
            .Cell(iRow, 3).Width = InchesToPoints(1.5)
            .Cell(iRow, 3).LeftPadding = PixelsToPoints(30)
            .Cell(iRow, 3).Shading.BackgroundPatternColorIndex = wdBrightGreen
            
            oRS.MoveNext
            
            If Not oRS.EOF Then
                Call .Rows.Add
            End If
            
            iRow = iRow + 1
                        
        Loop
    
    End With
            
End Sub

Sub MailMerge()
    Dim objWord As Word.Application
    Dim objDocument As Word.Document
    Dim objMailMerge As Word.MailMerge
    Dim cConnection As String
        
    cConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Password="""";" & _
                  "User ID=Admin;" & _
                  "Data Source=C:\BookCode\SampleDatabase.mdb;" & _
                  "Mode=Read;" & _
                  "Extended Properties="""";" & _
                  "Jet OLEDB:System database="""";" & _
                  "Jet OLEDB:Registry Path="""";" & _
                  "Jet OLEDB:Database Password="""";" & _
                  "Jet OLEDB:Engine Type=5;" & _
                  "Jet OLEDB:"
        
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True
    
    Set objDocument = objWord.Documents.Open(App.Path & "\mailmerge.doc")
     
    Set objMailMerge = objDocument.MailMerge
        
    Call objMailMerge.OpenDataSource( _
        "C:\BookCode\SampleDatabase.mdb", ConfirmConversions:=False, ReadOnly:= _
        True, LinkToSource:=True, AddToRecentFiles:=False, PasswordDocument:="", _
         PasswordTemplate:="", WritePasswordDocument:="", WritePasswordTemplate:= _
        "", Revert:=False, Format:=wdOpenFormatAuto, Connection:= _
        cConnection _
        , SQLStatement:="SELECT * FROM `Requester`", SQLStatement1:="", SubType:= _
        wdMergeSubTypeAccess)
        
    With objMailMerge
        .HighlightMergeFields = True
        .Destination = wdSendToNewDocument
        .Execute
    End With
                
End Sub

Sub MailingLabels()
    Dim objWord As Word.Application
    Dim objDocument As Word.Document
    Dim objAutoTextEntry As Word.AutoTextEntry
    Dim cConnection As String
        
    cConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Password="""";" & _
                  "User ID=Admin;" & _
                  "Data Source=C:\BookCode\SampleDatabase.mdb;" & _
                  "Mode=Read;" & _
                  "Extended Properties="""";" & _
                  "Jet OLEDB:System database="""";" & _
                  "Jet OLEDB:Registry Path="""";" & _
                  "Jet OLEDB:Database Password="""";" & _
                  "Jet OLEDB:Engine Type=5;" & _
                  "Jet OLEDB:"
           
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True
    
    Set objDocument = objWord.Documents.Add
    
    With objDocument.MailMerge
        
        'Temporarily set up the merge fields required
        With .Fields
            .Add objWord.Selection.Range, "FirstName"
            objWord.Selection.TypeText " "
            .Add objWord.Selection.Range, "LastName"
            objWord.Selection.TypeParagraph
            .Add objWord.Selection.Range, "Address1"
            objWord.Selection.TypeParagraph
            .Add objWord.Selection.Range, "City"
            objWord.Selection.TypeText ", "
            .Add objWord.Selection.Range, "State"
            objWord.Selection.TypeText " "
            .Add objWord.Selection.Range, "Zip"
        End With
        
        'Create an AutoText entry encapsulating the
        'merge fields for the label
        Set objAutoTextEntry = objWord.NormalTemplate. _
            AutoTextEntries.Add("Labels", objDocument.Content)
        
        'Remove the merges fields from the document as
        'objAutoTextEntry now contains what we need
        objDocument.Content.Delete
    
        Call .OpenDataSource( _
        "C:\BookCode\SampleDatabase.mdb", ConfirmConversions:=False, ReadOnly:= _
        True, LinkToSource:=True, AddToRecentFiles:=False, PasswordDocument:="", _
         PasswordTemplate:="", WritePasswordDocument:="", WritePasswordTemplate:= _
        "", Revert:=False, Format:=wdOpenFormatAuto, Connection:= _
        cConnection _
        , SQLStatement:="SELECT * FROM Requester", SQLStatement1:="", SubType:= _
        wdMergeSubTypeAccess)
                
        .MainDocumentType = wdMailingLabels
                
        Call objWord.MailingLabel.CreateNewDocument("5163", "", "Labels", wdPrinterManualFeed)
    
        .Destination = wdSendToNewDocument
        .Execute
    
        objAutoTextEntry.Delete
    
        objWord.NormalTemplate.Saved = True
    
    End With
    
End Sub

Sub ExcelReport()
    Dim oConn As New ADODB.Connection
    Dim oRS As ADODB.Recordset
    Dim objExcel As Excel.Application
    Dim objWB As Excel.Workbook
    Dim objWS As Excel.Worksheet
    Dim cSQL As String
    Dim cConnectString As String
    Dim cCompanyName As String
    Dim cGrandTotal As String
    Dim cFormula As String
    Dim iRow As Integer
    Dim iStartRow As Integer
    
    cConnectString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;" & _
                     "Persist Security Info=False;" & _
                     "Initial Catalog=EnterpriseReports;" & _
                     "Data Source=(local)"
                     
    oConn.ConnectionString = cConnectString
    oConn.Open
    
    cSQL = "SELECT Customers.CompanyName, Customers.ContactName, " & _
           "Orders.RequiredDate, Orders.Freight " & _
           "FROM { oj Northwind.dbo.Customers Customers " & _
           "INNER JOIN Northwind.dbo.Orders Orders ON " & _
           "Customers.CustomerID = Orders.CustomerID} " & _
           "ORDER BY Customers.CompanyName ASC"
           
    Set oRS = oConn.Execute(cSQL)
           
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    
    Set objWB = objExcel.Workbooks.Add
    
    Set objWS = objWB.Worksheets(1)
    
    iRow = 1
    cGrandTotal = "@SUM("
    
    With objWS
        .Cells(iRow, "A").Value = "Company"
        .Cells(iRow, "B").Value = "Contact"
        .Cells(iRow, "C").Value = "RequiredDate"
        .Cells(iRow, "D").Value = "Freight"
    
        .Cells(iRow, "A").Font.Bold = True
        .Cells(iRow, "B").Font.Bold = True
        .Cells(iRow, "C").Font.Bold = True
        .Cells(iRow, "D").Font.Bold = True
    End With
    
    iRow = iRow + 1
    
        
    Do While Not oRS.EOF
            
        If cCompanyName <> oRS("CompanyName") Then
                    
            If cCompanyName <> vbNullString Then
            
                cFormula = "@SUM(D" & iStartRow & ":D" & iRow - 1 & ")"
            
                objWS.Cells(iRow, "E").Font.Bold = True
                objWS.Cells(iRow, "E").Value = cFormula
                iRow = iRow + 1
                
                cGrandTotal = cGrandTotal & "E" & iRow - 1 & "+"
                
            End If
            
            iStartRow = iRow
            
            cCompanyName = oRS("CompanyName")
            
            objWS.Cells(iRow, "A").Value = cCompanyName
            objWS.Cells(iRow, "B").Value = oRS("ContactName")
            
            
        End If
        
        objWS.Cells(iRow, "C").Value = Format(oRS("RequiredDate"), "dd-mmm-yyyy")
        objWS.Cells(iRow, "D").Value = oRS("Freight")
                                    
        iRow = iRow + 1
    
        oRS.MoveNext
        
    Loop
    
    cGrandTotal = cGrandTotal & "E" & iRow & ")"
    cFormula = "@SUM(D" & iStartRow & ":D" & iRow - 1 & ")"
        
    With objWS
        .Cells(iRow, "E").Font.Bold = True
        .Cells(iRow, "E").Value = cFormula
        
        iRow = iRow + 1
        .Cells(iRow, "F").Font.Bold = True
        .Cells(iRow, "F").Value = cGrandTotal
        
        .Range("A1:F" & iRow).Columns.AutoFit
    End With
    
End Sub

Sub ExcelGraph()
     Dim oConn As New ADODB.Connection
     Dim oRS As ADODB.Recordset
     Dim objExcel As Excel.Application
     Dim objWB As Excel.Workbook
     Dim objWS As Excel.Worksheet
     Dim objChartObj As Excel.Chart
     Dim objSourceRange As Excel.Range
     Dim cSQL As String
     Dim cConnectString As String
     Dim X As Integer
     
     Screen.MousePointer = vbHourglass
     
     cConnectString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;" & _
                      "Persist Security Info=False;" & _
                      "Initial Catalog=Northwind;" & _
                      "Data Source=(local)"
                      
     oConn.ConnectionString = cConnectString
     oConn.Open
     
     cSQL = "SELECT MONTH(Orders.RequiredDate) AS MonthYr, " & _
            "SUM(Orders.Freight) AS FreightTotal " & _
            "FROM Orders " & _
            "WHERE Year(Orders.RequiredDate) = 1997 " & _
            "GROUP BY MONTH(Orders.RequiredDate) " & _
            "ORDER BY MONTH(Orders.RequiredDate)"
    
     Set oRS = oConn.Execute(cSQL)
     
     X = 1
    
     Set objExcel = CreateObject("Excel.Application")
     
     Set objWB = objExcel.Workbooks.Add
         
     Set objWS = objWB.Worksheets.Add
     
     Do While Not oRS.EOF
     
         objWS.Cells(X, 1) = Chr(39) & Format(oRS("MonthYr") & "/01", "mmmm")
         objWS.Cells(X, 2) = Val(oRS("FreightTotal"))
         
         X = X + 1
     
         oRS.MoveNext
     Loop
     
    
     ' Determine the size of the range and store it.
     Set objSourceRange = objWS.Range("A1:B" & X - 1)
     
     ' Create a new chart.
     Set objChartObj = objExcel.Charts.Add
     
     With objChartObj
     
        .ChartType = xlColumnClustered
        
        ' Set the range of the chart.
        .SetSourceData Source:=objSourceRange, PlotBy:=xlColumns
                   
        ' Specify that the chart is located on a new sheet.
        .Location Where:=xlLocationAsNewSheet
        
        ' Create and set the title; set title font.
        .HasTitle = True
        
        With .ChartTitle
           .Characters.Text = "Freight Charges by Month - 1997"
           .Characters.Font.Color = vbRed
           .Font.Size = 16
        End With
    
        ' Delete the legend.
        .HasLegend = False
        
        With .SeriesCollection(1)
           .ApplyDataLabels Type:=xlDataLabelsShowValue
           .DataLabels.NumberFormat = "#,##0"
        End With
        
        .Export App.Path & "\mychart.jpg", "JPEG"
     
     End With
    
     oRS.Close
     Set oRS = Nothing
     
     objExcel.Visible = True
     
     Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
    oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BookCode\SampleDatabase.mdb;Persist Security Info=False"
    oConn.Open
End Sub

Private Sub OutlookEmail()
    Dim objEMail As New EMail
    
    On Error GoTo EmailError
    
    With objEMail
    
        .StartOutlook
        
        .Subject = "Background Check Completed"
        
        Call .AddRecipient("joerecruiter@hiscompany.com")
        Call .AddRecipient("Ganz, Carl")
    
        
        If Not .ResolveAll Then
            Err.Raise 512 + vbObjectError
        End If
        
        .Body = "Congratulations, your applicant passed the background check. " & _
                "Please have him fill out the attached W2 form."
                
        Call .AddAttachments("c:\w2.doc", 0, olByValue, "Link")
        
        .Save
        .Send
        .CloseOutlook
            
    End With
    
EmailError:
    
    Select Case Err.Number
    
        Case 512 + vbObjectError
            MsgBox "Some of the email recipients were not found in the contact file."
            Exit Sub
    
    End Select
        
End Sub

Public Function LotusNotesEmail(ByVal cRecipient As String, ByVal cCopyTo As String, _
                          ByVal cSubject As String, ByVal cBody As String, _
                          Optional vntAttachment As Variant, _
                          Optional vntSaveMessageOnSend As Variant) As Boolean
                          
    Dim objNotes As Object
    Dim objNotesDB As Object
    Dim objNotesMailDoc As Object
    Dim objNotesRichText As Object
    Dim objNotesEmbedded As Object
    Dim objNotesDir As Object
    Dim cPassword As String
    Dim cMsg As String
    Dim bResult As Boolean
    Dim X As Integer
    Dim aRecipient() As Variant
    Dim aAttachment() As Variant
    Dim ValueSALT As String
    
    
    Const ERROR_NO_LOTUS_NOTES_PASSWORD = 1
    
    ValueSALT = "SALTValue"
    
    bResult = False
    
    If cRecipient = vbNullString And _
       cCopyTo = vbNullString Then
       
        LotusNotesEmail = bResult
       
        Exit Function
        
    End If
    
    On Error GoTo ErrorHandler
    
    If IsMissing(vntSaveMessageOnSend) Then
        vntSaveMessageOnSend = True
    End If
    
    aRecipient = Parse2Array(cRecipient, ";")
    
    
    Set objNotes = GetObject("", "Lotus.NotesSession")
            
    cPassword = GetSetting("LotusNotes", "AppInfo", "LotusNotesPassword")

    cPassword = DecryptString(cPassword, ValueSALT)

    If cPassword = vbNullString Then
        Err.Raise ERROR_NO_LOTUS_NOTES_PASSWORD
    End If
    
    'In order for this line of code not to prompt for a password
    'every time it is invoked, the local installation of Lotus Notes
    'needs to have the "Don't prompt for a password from
    'other Notes-base programs" option checked under File|Tools|User ID
    Call objNotes.Initialize(cPassword)
    
        
    Call SaveSetting("LotusNotes", "AppInfo", "LotusNotesPassword", _
    EncryptString(cPassword, ValueSALT))
                
    Set objNotesDir = objNotes.GetDbDirectory("")
    Set objNotesDB = objNotesDir.OpenMailDatabase
    
    Set objNotesMailDoc = objNotesDB.CreateDocument
    
    With objNotesMailDoc
        
        .SaveMessageOnSend = vntSaveMessageOnSend
        
        Call .AppendItemValue("Form", "Memo")
        Call .AppendItemValue("Subject", cSubject)
        Call .AppendItemValue("SendTo", aRecipient)
        
        If cCopyTo <> vbNullString Then
            Call .AppendItemValue("CopyTo", cCopyTo)
        End If
    
        If .HasItem("body") Then
            Set objNotesRichText = .GetFirstItem("body")
        Else
            Set objNotesRichText = .CreateRichTextItem("Body")
        End If
    
    End With
        
    Call objNotesRichText.AppendText(cBody)
    Call objNotesRichText.AddNewLine(2)
    
    If Not IsMissing(vntAttachment) Then
    
        aAttachment = Parse2Array(vntAttachment, ";")
    
        For X = 0 To UBound(aAttachment)
            Set objNotesEmbedded = objNotesRichText.EmbedObject(1454, _
        "", aAttachment(X))
            objNotesMailDoc.CreateRichTextItem ("Attachment" & X)
        Next X
        
    End If
    
    Call objNotesMailDoc.Save(True, True)
    Call objNotesMailDoc.Send(True)
    
    bResult = True
    
    Set objNotes = Nothing
    Set objNotesDB = Nothing
    Set objNotesMailDoc = Nothing
    Set objNotesRichText = Nothing
    Set objNotesEmbedded = Nothing
    Set objNotesDir = Nothing

    LotusNotesEmail = bResult

Exit Function

ErrorHandler:

    Select Case Err.Number
    
        Case ERROR_NO_LOTUS_NOTES_PASSWORD
            cMsg = "There is no Lotus Notes password on file with this application. " & _
            vbCrLf & vbCrLf & _
                   "Please enter the correct password and press OK to save and continue."
            
            cPassword = InputBox(cMsg, "Notes Password Invalid")
                        
            Resume Next
    
    End Select

    LotusNotesEmail = bResult
    
End Function

Function DecryptString(cData1 As String, cData2 As String) As String
    DecryptString = cData1
End Function

Function EncryptString(cData1 As String, cData2 As String) As String
    EncryptString = cData1
End Function

Function Parse2Array(ByVal cData As String, cSep As String) As Variant()
    Dim aResult() As Variant
    Dim lCnt As Long
    Dim iSepPos As Integer
    Dim cName As String
    
    cData = Trim(cData)
    
    If Mid$(cData, 1, 1) = "(" Then
        cData = Mid$(cData, 2)
    
        If Mid$(cData, Len(cData), 1) = ")" Then
            cData = Mid$(cData, 1, Len(cData) - 1)
        End If
    End If
    

   
    Do While True
    
        iSepPos = InStr(cData, cSep)
    
        If iSepPos = 0 Then
            cName = cData
            cData = vbNullString
        Else
            cName = Mid$(cData, 1, iSepPos - 1)
        End If
                
        ReDim Preserve aResult(lCnt)
        
        aResult(lCnt) = Trim$(cName)
        
        lCnt = lCnt + 1
        
        If cData = vbNullString Then
            Exit Do
        End If
        
        cData = Mid$(cData, iSepPos + 1)
    
    Loop

    Parse2Array = aResult
    
End Function

Sub PivotTable()
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
