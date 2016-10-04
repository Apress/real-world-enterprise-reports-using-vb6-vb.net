VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReportCriteria 
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   5250
   Begin VB.ComboBox cboComboBox 
      Height          =   315
      Index           =   0
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSComctlLib.ListView lstListBox 
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2990
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CheckBox chkCriteriaPage 
      Caption         =   "Print Criteria Page"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin EditLib.fpText txtTextBox 
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   503
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   1
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
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
   Begin EditLib.fpDateTime txtDateFrom 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
      _Version        =   196608
      _ExtentX        =   2355
      _ExtentY        =   503
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   3
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "00000000"
      DateMin         =   "00000000"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "00000000"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   1
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime txtDateTo 
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
      _Version        =   196608
      _ExtentX        =   2355
      _ExtentY        =   503
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   3
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "00000000"
      DateMin         =   "00000000"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "00000000"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   1
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label lblWidthChecker 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblDateRange 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
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
            cSQL = GetDateSQL("createdate", cDateFrom, cDateTo)
            cSQL = cSQL & GetListBoxSQL(SELECT_PRODUCT, "productid")
            cSQL = cSQL & GetListBoxSQL(SELECT_SOURCE, "sourceid")
            cSQL = cSQL & GetComboBoxSQL(SELECT_DEPARTMENT, "deptid")
            cSQL = cSQL & GetCheckBoxSQL(SELECT_COMPLEX_QUESTION, "questiontype")
            
            cCriteria = GetDateCriteria("Create Date", cDateFrom, cDateTo)
            cCriteria = cCriteria & GetListBoxCriteria(SELECT_PRODUCT, "Product")
            cCriteria = cCriteria & GetListBoxCriteria(SELECT_SOURCE, "Source")
            cCriteria = cCriteria & GetComboBoxCriteria(SELECT_DEPARTMENT, "Department")
            cCriteria = cCriteria & GetCheckBoxCriteria(SELECT_COMPLEX_QUESTION, "Complex Question")
        
        Case "DurationByProductRpt"
        
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
     
        .MultiSelect = MultiSelectExtended
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

    With lblComboBox(iIndex)
        .Caption = cCaption
        .AutoSize = True
        .Top = iTop
        .Left = iLeft
        .Width = lblComboBox(iIndex).Width
        .Visible = True
    End With
    
    With cboComboBox(iIndex)
        .Top = iTop
        .Left = lblComboBox(iIndex).Width + 200
        .Width = iWidth
        .Visible = True
    End With
    
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
        
    cList = ParseIt(Index, True, False, 0)
    
    cSQL = cColumn & " IN " & cList & " AND "
                           
    GetListBoxSQL = cSQL
    
End Function

Function GetListBoxCriteria(Index As Integer, cDescr As String) As String
    Dim cResult As String
    Dim cNames As String
        
    cNames = ParseIt(Index, True, False, 1)
    
    cResult = cResult & "the " & cDescr & _
        " is among " & cNames & " and " & vbCrLf

    GetListBoxCriteria = cResult
    
End Function

Function GetComboBoxSQL(Index As Integer, cColumn As String) As String
    Dim cSQL As String
    Dim lChoiceID As Long
    
    lChoiceID = Mid$(cboComboBox(Index), InStr(cboComboBox(Index), vbTab) + 1)
        
    If lChoiceID > 0 Then
    
        cSQL = cColumn & " = " & lChoiceID & " AND "
                
    End If
           
    GetComboBoxSQL = cSQL
    
End Function

Function GetComboBoxCriteria(Index As Integer, cDescr As String) As String
    Dim cResult As String
    Dim cChoice As String
        
    cChoice = Trim$(Mid$(cboComboBox(Index), 1, InStr(cboComboBox(Index), vbTab) - 1))
    
    If cChoice <> vbNullString Then
    
        cResult = "the " & cDescr & " is " & cChoice & " AND "
                
    End If
           
    GetComboBoxCriteria = cResult
    
End Function

Function GetCheckBoxSQL(Index As Integer, cColumn As String) As String
    Dim cSQL As String
    Dim iChoice As Integer
    
    iChoice = chkCheckBox(Index)
    
    cSQL = cColumn & " = " & IIf(iChoice = vbChecked, BOOL_TRUE, BOOL_FALSE) & " AND "
                
    GetCheckBoxSQL = cSQL
    
End Function

Function GetCheckBoxCriteria(Index As Integer, cDescr As String) As String
    Dim cResult As String
    Dim cChoice As String
    
    cChoice = IIf(chkCheckBox(Index) = vbChecked, "True", "False")
    
    cResult = "the " & cDescr & " is " & cChoice & " AND "
                
    GetCheckBoxCriteria = cResult
    
End Function

Function ParseIt(Index As Integer, bTagged As Boolean, bQuotes As Boolean, iCol As Variant) As String
    Dim cResult As String
    Dim iCount As Integer
    Dim x As Integer
    
    iCount = lstListBox(Index).ListItems.Count
    
    For x = 1 To iCount
    
        If lstListBox(Index).ListItems(x).Checked Then
            If iCol = 1 Then
                cResult = cResult & lstListBox(Index).ListItems(x).Text & ","
            Else
                cResult = cResult & lstListBox(Index).ListItems(x).SubItems(1) & ","
            End If
        End If
        
    Next x
    
    If cResult <> vbNullString Then
        cResult = Mid$(cResult, 1, Len(cResult) - Len(","))
        cResult = "(" & cResult & ")"
    End If
    
    ParseIt = cResult
    
End Function

Sub LoadListBox(oConn As ADODB.Connection, oListBox As Object, cTable As String, cID As String, cDescr1 As String)
    Dim oRS As Recordset
    Dim cSQL As String
    Dim itmX As ListItem
    
    cSQL = "SELECT * " & _
           "FROM " & cTable & _
           " ORDER BY " & cDescr1

    Set oRS = RunCommand(oConn, cSQL, adCmdText, adOpenForwardOnly, False)

    Do While Not oRS.EOF

        Set itmX = oListBox.ListItems.Add()
        itmX.Text = oRS(cDescr1)
        itmX.SubItems(1) = oRS(cID)

        oRS.MoveNext
        
    Loop

    oRS.Close
    Set oRS = Nothing

End Sub

Sub LoadComboBox(oConn As ADODB.Connection, oCombobox As Object, cTable As String, cID As String, cDescr1 As String)
    Dim oRS As Recordset
    Dim cSQL As String
    Dim itmX As ListItem
    
    cSQL = "SELECT * " & _
           "FROM " & cTable & _
           " ORDER BY " & cDescr1

    Set oRS = RunCommand(oConn, cSQL, adCmdText, adOpenForwardOnly, False)

    Do While Not oRS.EOF

        oCombobox.AddItem oRS(cDescr1) & Space(40) & vbTab & oRS(cID)

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

Private Sub cmdListBox_Click(Index As Integer)
    Dim iCount As Integer
    Dim x As Integer
    
    iCount = lstListBox(Index).ListItems.Count
    
    For x = 1 To iCount
        lstListBox(Index).ListItems(x).Checked = False
    Next x
    
End Sub

Sub FormCenter(frmMainMenu As Object, frmCurrent As Object)
   frmCurrent.Move frmMainMenu.ScaleWidth / 2 - frmCurrent.Width / 2, frmMainMenu.ScaleHeight / 2 - frmCurrent.Height / 2
End Sub

