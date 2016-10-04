VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReportCriteria 
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   5250
   Begin MSComCtl2.DTPicker txtDateFrom 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   37432
   End
   Begin VB.PictureBox DTPicker1 
      Height          =   285
      Left            =   -5280
      ScaleHeight     =   225
      ScaleWidth      =   1155
      TabIndex        =   14
      Top             =   -3120
      Width           =   1215
   End
   Begin VB.ComboBox cboComboBox 
      Height          =   315
      Index           =   0
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSComctlLib.ListView lstListBox 
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3201
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.TextBox txtTextBox 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox chkCriteriaPage 
      Caption         =   "Print Criteria Page"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox chkCheckBox 
      Caption         =   "Check1"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdListBox 
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker txtDateTo 
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   37432
   End
   Begin VB.Label lblWidthChecker 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblDateRange 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblTextBox 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblComboBox 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblListBox 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmReportCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SELECT_PRODUCT = 1
Const SELECT_SOURCE = 2

Const SELECT_ENTER_DATE = 1
Const SELECT_RECEIVE_DATE = 2

Const SELECT_FIRST_NAME = 1
Const SELECT_LAST_NAME = 2

Const SELECT_DEPARTMENT = 1

Const SELECT_COMPLEX_QUESTION = 1
     
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdListBox_Click(Index As Integer)
    Dim iCount As Integer
    Dim x As Integer
    
    iCount = lstListBox(Index).ListItems.Count
    
    For x = 1 To iCount
        lstListBox(Index).ListItems(x).Checked = False
    Next x
End Sub

Private Sub cmdOK_Click()
    Dim iDateCnt As Integer
    Dim cDateFrom As String
    Dim cDateTo As String
    Dim cDescr As String
    Dim cSQL As String
    Dim cCriteria As String
    Dim x As Integer
    
    iDateCnt = txtDateFrom.Count - 1

    For x = 0 To iDateCnt
    
        cDateFrom = txtDateFrom(x)
        cDateTo = txtDateTo(x)
        
        If cDateFrom <> vbNullString Or cDateTo <> vbNullString Then
            
            If Not IsDate(cDateFrom) Or Not IsDate(cDateTo) Then
                
                MsgBox "You must enter both a valid 'from' and a 'to' date.", vbOKOnly
            
                Exit Sub
            
            End If
            
        End If
                
                
        If IsDate(cDateFrom) And IsDate(cDateTo) Then
        
            If DateValue(cDateFrom) > DateValue(cDateTo) Then
            
                MsgBox "The 'from date' cannot be greater than the 'to date'.", vbOKOnly
                
                txtDateFrom(x).SetFocus
                
                Exit Sub
                            
            End If
        
        End If
        
    Next x
    
        
    Select Case objReport.Report
    
        Case "DurationBySourceRpt"
            cSQL = GetDateSQL("createdate", txtDateFrom(1), txtDateTo(1))
            cSQL = cSQL & GetListBoxSQL(SELECT_PRODUCT, "productid")
            cSQL = cSQL & GetListBoxSQL(SELECT_SOURCE, "sourceid")
            cSQL = cSQL & GetComboBoxSQL(SELECT_DEPARTMENT, "deptid")
            cSQL = cSQL & GetCheckBoxSQL(SELECT_COMPLEX_QUESTION, "questiontype")
            
            cCriteria = GetDateCriteria("Create Date", txtDateFrom(1), txtDateTo(1))
            cCriteria = cCriteria & GetListBoxCriteria(SELECT_PRODUCT, "Product")
            cCriteria = cCriteria & GetListBoxCriteria(SELECT_SOURCE, "Source")
            cCriteria = cCriteria & GetComboBoxCriteria(SELECT_DEPARTMENT, "Department")
            cCriteria = cCriteria & GetCheckBoxCriteria(SELECT_COMPLEX_QUESTION, "Complex Question")
        
        Case "DurationByProductRpt"
            cSQL = GetDateSQL("createdate", txtDateFrom(1), txtDateTo(1))
            cSQL = cSQL & GetDateSQL("receivedate", txtDateFrom(2), txtDateTo(2))
            cSQL = cSQL & GetListBoxSQL(SELECT_PRODUCT, "productid")
            
            cCriteria = GetDateCriteria("Create Date", txtDateFrom(1), txtDateTo(1))
            cCriteria = cCriteria & GetDateCriteria("Receive Date", txtDateFrom(2), txtDateTo(2))
            cCriteria = cCriteria & GetListBoxCriteria(SELECT_PRODUCT, "Product")
            
    End Select
    
    If cSQL <> vbNullString Then
        cSQL = Mid$(cSQL, 1, Len(cSQL) - Len(" AND "))
    End If

    If cCriteria <> vbNullString Then
        cCriteria = Mid$(cCriteria, 1, Len(cCriteria) - Len(" AND "))
    End If

    With objReport
        .PrintCriteriaPage = IIf(chkCriteriaPage = vbChecked, True, False)
        .ReportCriteria = cCriteria
        .ReportSQL = cSQL
        .ShowPrint
    End With
    
End Sub

Private Sub Form_Load()

    Me.Caption = "Criteria for " & objReport.ReportTitle

    Select Case objReport.Report
    
        Case "DurationBySourceRpt"
            Call ShowDates(SELECT_ENTER_DATE, 120, 100, "Enter Date Range:")
            Call ShowListBox(SELECT_PRODUCT, 1400, 100, 3000, "Product")
            Call ShowListBox(SELECT_SOURCE, 1400, 3500, 3000, "Source")
            Call ShowComboBox(SELECT_DEPARTMENT, 4300, 100, 3000, "Department")
            Call ShowCheckBox(SELECT_COMPLEX_QUESTION, 4300, 4500, "Complex Question")
            
            Call LoadListBox(oConn, lstListBox(SELECT_PRODUCT), "product", "ID", "Descr")
            Call LoadListBox(oConn, lstListBox(SELECT_SOURCE), "source", "ID", "Descr")

            Call LoadComboBox(oConn, cboComboBox(SELECT_DEPARTMENT), "department", "ID", "Descr")
            
            Me.Width = 7000
            Me.Height = 6000
        
        
        Case "DurationByProductRpt"
            Call ShowDates(SELECT_ENTER_DATE, 120, 100, "Enter Date Range:")
            Call ShowDates(SELECT_RECEIVE_DATE, 700, 100, "Enter Receive Range:")
            Call ShowListBox(SELECT_PRODUCT, 1500, 100, 3000, "Product")
            
            Call LoadListBox(oConn, lstListBox(SELECT_PRODUCT), "product", "ID", "Descr")
            
            Me.Width = 7000
            Me.Height = 4500
        
    End Select
        
    Call PositionButtons
    
    Call FormCenter(frmMainMenu, Me)

End Sub

Sub ShowDates(iIndex As Integer, iTop As Integer, iLeft As Integer, cCaption As String)
    Load lblDateRange(iIndex)
    Load txtDateFrom(iIndex)
    Load txtDateTo(iIndex)
    
    lblDateRange(iIndex).Top = iTop
    txtDateFrom(iIndex).Top = iTop
    txtDateTo(iIndex).Top = iTop
    
    lblDateRange(iIndex).Left = iLeft
    txtDateFrom(iIndex).Left = lblDateRange(iIndex).Width + iLeft + 200
    txtDateTo(iIndex).Left = txtDateFrom(iIndex).Left + txtDateFrom(iIndex).Width + 200
    
    lblDateRange(iIndex) = cCaption
    
    lblDateRange(iIndex).Visible = True
    txtDateFrom(iIndex).Visible = True
    txtDateTo(iIndex).Visible = True
    
End Sub

Sub ShowListBox(iIndex As Integer, iTop As Integer, _
    iLeft As Integer, iWidth As Integer, cCaption As String)
    Dim clmX As ColumnHeader
    
    Load lblListBox(iIndex)
    Load lstListBox(iIndex)
    Load cmdListBox(iIndex)
    
    With lblListBox(iIndex)
        .Top = iTop
        .Left = iLeft
        .Width = iWidth
        .Caption = "Select " & cCaption
        .Visible = True
    End With
    
    With lstListBox(iIndex)
        Set clmX = lstListBox(iIndex).ColumnHeaders.Add()
        clmX.Width = iWidth
        Set clmX = lstListBox(iIndex).ColumnHeaders.Add()
        clmX.Width = 0
     
        '.MultiSelect = MultiSelectExtended
        .Top = lblListBox(iIndex).Top + _
               lblListBox(iIndex).Height + 100
        .Left = iLeft
        .Width = iWidth
        .View = lvwReport
        .Checkboxes = True
        .Visible = True
    End With
    
    With cmdListBox(iIndex)
        .Left = iLeft
        .Width = iWidth
        .Top = lstListBox(iIndex).Top + _
                lstListBox(iIndex).Height + 100
        .Caption = "Clear All " & cCaption & " Selections"
        .Visible = True
    End With
    
End Sub


Sub ShowComboBox(iIndex As Integer, iTop As Integer, iLeft As Integer, iWidth As Integer, cCaption As String)
   
    Load lblComboBox(iIndex)
    Load cboComboBox(iIndex)
    
    lblComboBox(iIndex) = cCaption
    lblComboBox(iIndex).AutoSize = True
    
    cboComboBox(iIndex).Top = iTop
    cboComboBox(iIndex).Left = lblComboBox(iIndex).Width + 200
    cboComboBox(iIndex).Width = iWidth

    lblComboBox(iIndex).Top = iTop
    lblComboBox(iIndex).Left = iLeft
    lblComboBox(iIndex).Width = lblComboBox(iIndex).Width
    lblComboBox(iIndex) = cCaption
    
    lblComboBox(iIndex).Visible = True
    cboComboBox(iIndex).Visible = True
    
End Sub

Sub ShowCheckBox(iIndex As Integer, iTop As Integer, iLeft As Integer, cCaption As String)
    Dim iTextWidth As Integer

    Load chkCheckBox(iIndex)

    Set lblWidthChecker.Font = chkCheckBox(iIndex).Font
    lblWidthChecker = cCaption
    lblWidthChecker.AutoSize = True
    
    iTextWidth = lblWidthChecker.Width + 300
    
    With chkCheckBox(iIndex)
        .Top = iTop
        .Left = iLeft
        .Width = iTextWidth
        .Caption = cCaption
        .Visible = True
    End With
    
End Sub

Sub PositionButtons()

    cmdOK.Left = Me.Width - cmdOK.Width - 200
    cmdCancel.Left = Me.Width - cmdCancel.Width - 200
    cmdHelp.Left = Me.Width - cmdHelp.Width - 200
    
End Sub

Function GetDateSQL(cFieldName As String, cDateFrom As String, cDateTo As String) As String
    Dim cSQL As String
        
    If IsDate(cDateFrom) And IsDate(cDateTo) Then
        
        cSQL = cFieldName & " BETWEEN " & Chr(39) & cDateFrom & Chr(39) & " AND " & Chr(39) & cDateTo & Chr(39) & " AND "
               
    End If
    
    GetDateSQL = cSQL
    
End Function

Function GetDateCriteria(cDescr As String, cDateFrom As String, cDateTo As String) As String
    Dim cResult As String
        
    If IsDate(cDateFrom) And IsDate(cDateTo) Then
        
        cResult = cResult & "the " & cDescr & " is between " & cDateFrom & " and " & cDateTo & " and " & vbCrLf
                                
    End If
    
    GetDateCriteria = cResult
    
End Function

Function GetListBoxSQL(Index As Integer, cColumn As String) As String
    Dim cSQL As String
    Dim cList As String
        
    cList = ParseIt(lstListBox(Index), False, 0)
    
    cSQL = cColumn & " IN " & cList & " AND "
                           
    GetListBoxSQL = cSQL
    
End Function

Function GetListBoxCriteria(Index As Integer, cDescr As String) As String
    Dim cResult As String
    Dim cNames As String
        
    cNames = ParseIt(lstListBox(Index), False, 1)
    
    cResult = cResult & "the " & cDescr & _
        " is among " & cNames & " and " & vbCrLf

    GetListBoxCriteria = cResult
    
End Function

Function GetComboBoxSQL(Index As Integer, cColumn As String) As String
    Dim cSQL As String
    Dim lChoiceID As Long
    
    If cboComboBox(Index) <> vbNullString Then
    
        lChoiceID = Mid$(cboComboBox(Index), InStr(cboComboBox(Index), vbTab) + 1)
            
        If lChoiceID > 0 Then
        
            cSQL = cColumn & " = " & lChoiceID & " AND "
                    
        End If
    
    End If
           
    GetComboBoxSQL = cSQL
    
End Function

Function GetComboBoxCriteria(Index As Integer, cDescr As String) As String
    Dim cResult As String
    Dim cChoice As String
    
    If cboComboBox(Index) <> vbNullString Then
    
        cChoice = Trim$(Mid$(cboComboBox(Index), 1, InStr(cboComboBox(Index), vbTab) - 1))
        
        If cChoice <> vbNullString Then
        
            cResult = "the " & cDescr & " is " & cChoice & " and "
                    
        End If
    
    End If
           
    GetComboBoxCriteria = cResult
    
End Function

Function GetCheckBoxSQL(Index As Integer, cColumn As String) As String
    Dim cSQL As String
    Dim iChoice As Integer
    
    iChoice = chkCheckBox(Index)
        
    If iChoice <> vbGrayed Then
    
        cSQL = cColumn & " = " & IIf(iChoice = vbChecked, BOOL_TRUE, BOOL_FALSE) & " AND "
                
    End If
           
    GetCheckBoxSQL = cSQL
    
End Function

Function GetCheckBoxCriteria(Index As Integer, cDescr As String) As String
    Dim cResult As String
    Dim iChoice As Integer
    
    iChoice = chkCheckBox(Index)
    
    If iChoice <> vbGrayed Then
    
        cResult = "the " & cDescr & " is " & IIf(iChoice = vbChecked, "True", "False") & " AND "
                
    End If
           
    GetCheckBoxCriteria = cResult
    
End Function

Function ParseIt(objList As Object, bQuotes As Boolean, iCol As Integer) As String
    Dim cResult As String
    Dim cQuotes As String
    Dim iCount As Integer
    Dim x As Integer
    
    If bQuotes Then
        cQuotes = Chr(39)
    End If
    
    iCount = objList.ListItems.Count
    
    For x = 1 To iCount
    
        If objList.ListItems(x).Checked Then
            If iCol = 0 Then
                cResult = cResult & cQuotes & objList.ListItems(x).SubItems(1) & cQuotes & ","
            Else
                cResult = cResult & cQuotes & objList.ListItems(x).Text & cQuotes & ","
            End If
        End If
        
    Next x
    
    If cResult <> vbNullString Then
        cResult = Mid$(cResult, 1, Len(cResult) - Len(","))
        cResult = "(" & cResult & ")"
    End If
    
    ParseIt = cResult
    
End Function

Sub LoadListBox(oConn As ADODB.Connection, oList As Object, cTable As String, cID As String, cDescr1 As String)
    Dim itmX As ListItem
    Dim oRS As Recordset
    Dim cSQL As String

    cSQL = "SELECT * " & _
           "FROM " & cTable & _
           " ORDER BY " & cDescr1

    Set oRS = RunCommand(oConn, cSQL, adCmdText, adOpenForwardOnly, False)
      
    Do While Not oRS.EOF
        
        Set itmX = oList.ListItems.Add()
        itmX.Text = oRS(cDescr1)
        itmX.SubItems(1) = oRS(cID)
    
        oRS.MoveNext
        
    Loop

    oRS.Close
    Set oRS = Nothing

End Sub

Sub LoadComboBox(oConn As ADODB.Connection, oCombo As Object, cTable As String, cID As String, cDescr1 As String)
    Dim oRS As Recordset
    Dim cSQL As String

    cSQL = "SELECT * " & _
           "FROM " & cTable & _
           " ORDER BY " & cDescr1

    Set oRS = RunCommand(oConn, cSQL, adCmdText, adOpenForwardOnly, False)
      
    Do While Not oRS.EOF
        
        oCombo.AddItem oRS(cDescr1) & Space(40) & vbTab & oRS(cID)
    
        oRS.MoveNext
        
    Loop

    oRS.Close
    Set oRS = Nothing

End Sub

Function RunCommand(oConn As ADODB.Connection, cSQL As String, iCommandType As Integer, iCursorType As Integer, Optional vntColumn As Variant, Optional vntID As Variant) As ADODB.Recordset
    Dim oRS As New ADODB.Recordset
    Dim oCmd As New ADODB.Command

    oConn.CursorLocation = adUseClient

    With oCmd
    
        Set .ActiveConnection = oConn
        .CommandText = cSQL
        .CommandType = iCommandType
        .CommandTimeout = 1200
        
        If iCommandType = adCmdStoredProc And _
            Not IsMissing(vntColumn) And _
            Not IsMissing(vntID) Then
        
            If IsNumeric(vntID) Then
                oCmd.Parameters.Append oCmd.CreateParameter(vntColumn, adInteger, adParamInput, 9, vntID)
            Else
                oCmd.Parameters.Append oCmd.CreateParameter(vntColumn, adVarChar, adParamInput, 50, vntID)
            End If
            
        End If
    
    End With
    
    With oRS
        .CursorType = iCursorType
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .Open oCmd
    End With

    Set RunCommand = oRS

End Function

Sub FormCenter(frmMainMenu As Object, frmCurrent As Object)
   frmCurrent.Move frmMainMenu.ScaleWidth / 2 - frmCurrent.Width / 2, frmMainMenu.ScaleHeight / 2 - frmCurrent.Height / 2
End Sub
