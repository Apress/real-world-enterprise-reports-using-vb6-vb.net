Imports System.Text

Public Class frmReportCriteria
    Inherits System.Windows.Forms.Form

    Dim objListBoxLabelColl As New Collection()
    Dim objListBoxColl As New Collection()
    Dim objListBoxButtonColl As New Collection()

    Dim objComboBoxLabelColl As New Collection()
    Dim objComboBoxColl As New Collection()

    Dim objDateRangeLabelColl As New Collection()
    Dim objDateFromColl As New Collection()
    Dim objDateToColl As New Collection()

    Dim objCheckBoxColl As New Collection()

    Const SELECT_PRODUCT = 1
    Const SELECT_SOURCE = 2

    Const SELECT_ENTER_DATE = 1
    Const SELECT_RECEIVE_DATE = 2

    Const SELECT_FIRST_NAME = 1
    Const SELECT_LAST_NAME = 2

    Const SELECT_DEPARTMENT = 1

    Const SELECT_COMPLEX_QUESTION = 1

    Const BOOL_TRUE = -1
    Const BOOL_FALSE = 0

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(560, 24)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(80, 32)
        Me.cmdOK.TabIndex = 0
        Me.cmdOK.Text = "&OK"
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(560, 64)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(80, 32)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "Cancel"
        '
        'frmReportCriteria
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Silver
        Me.ClientSize = New System.Drawing.Size(648, 445)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOK})
        Me.Name = "frmReportCriteria"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim oConn As New OleDb.OleDbConnection()
        Dim cConnectString As String

        cConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BookCode\SampleDatabase.mdb"

        oConn.ConnectionString = cConnectString
        oConn.Open()

        SuspendLayout()

        ShowDateRange(SELECT_ENTER_DATE, 8, 32, 20, "Enter Date")
        ShowListBox(SELECT_PRODUCT, 8, 82, 180, 20, "Product")
        ShowListBox(SELECT_SOURCE, 300, 82, 180, 20, "Source")
        ShowComboBox(SELECT_DEPARTMENT, 8, 340, 180, 20, "Department")
        ShowCheckBox(SELECT_COMPLEX_QUESTION, 300, 340, 20, "Complex Question")

        LoadTable(oConn, objListBoxColl(SELECT_PRODUCT), "Product", "ID", "Descr")
        LoadTable(oConn, objListBoxColl(SELECT_SOURCE), "Source", "ID", "Descr")
        LoadTable(oConn, objComboBoxColl(SELECT_DEPARTMENT), "Department", "ID", "Descr")

        Me.Height() = 400
        Me.Width() = 600

        PositionButtons()

        ResumeLayout()

    End Sub

    Sub PositionButtons()
        cmdOK.Left() = Me.Width - cmdOK.Width - 10
        cmdOK.Top() = 10

        cmdCancel.Left() = cmdOK.Left
        cmdCancel.Top() = cmdOK.Top + cmdOK.Height + 10

    End Sub

    Sub LoadTable(ByVal oConn As OleDb.OleDbConnection, ByRef oControl As Object, _
        ByVal cTable As String, ByVal cID As String, ByVal cDescr As String)

        Dim oDS As New DataSet()
        Dim oDA As OleDb.OleDbDataAdapter
        Dim cSQL As String

        cSQL = "SELECT * " & _
               "FROM " & cTable & _
               " ORDER BY " & cDescr

        oDA = New OleDb.OleDbDataAdapter(cSQL, oConn)
        oDA.Fill(oDS, cTable)

        With oControl
            .DataSource = oDS.Tables(0)
            .DisplayMember = cDescr
            .ValueMember = cID
        End With

    End Sub

    Private Sub ShowListBox(ByVal iIndex As Short, ByVal iLeft As Short, ByVal iTop As Short, _
        ByVal iWidth As Short, ByVal iHeight As Short, ByVal cCaption As String)

        Call AddDynamicListBoxLabel(iIndex, iLeft, iTop, iWidth, iHeight, cCaption)

        Call AddDynamicListBox(iIndex, iLeft, objListBoxLabelColl(iIndex).Top + _
            objListBoxLabelColl(iIndex).Height + 5, iWidth, 180)

        Call AddDynamicListBoxButton(iIndex, iLeft, objListBoxColl(iIndex).Top + _
            objListBoxColl(iIndex).Height + 5, iWidth, iHeight, cCaption)

    End Sub

    Private Sub AddDynamicListBoxLabel(ByVal iIndex As Short, ByVal iLeft As Short, _
        ByVal iTop As Short, ByVal iWidth As Short, ByVal iHeight As Short, ByVal cCaption As String)

        objListBoxLabelColl.Add(New Label())

        With objListBoxLabelColl(iIndex)
            .Name = "ListBoxLabel" & iIndex
            .Size = New Size(iWidth, iHeight)
            .Location = New Point(iLeft, iTop)
            .Text = cCaption
        End With

        Me.Controls.AddRange(New System.Windows.Forms.Control() {objListBoxLabelColl(iIndex)})

    End Sub

    Private Sub AddDynamicListBox(ByVal iIndex As Short, ByVal iLeft As Short, _
        ByVal iTop As Short, ByVal iWidth As Short, ByVal iHeight As Short)

        objListBoxColl.Add(New ListBox())

        With objListBoxColl(iIndex)
            .Name = "ListBox" & iIndex
            .Size = New Size(iWidth, iHeight)
            .Location = New Point(iLeft, iTop)
            .SelectionMode = SelectionMode.MultiExtended
        End With

        Me.Controls.AddRange(New System.Windows.Forms.Control() {objListBoxColl(iIndex)})

    End Sub

    Private Sub AddDynamicListBoxButton(ByVal iIndex As Short, ByVal iLeft As Short, _
        ByVal iTop As Short, ByVal iWidth As Short, ByVal iHeight As Short, ByVal cCaption As String)
        Dim objButton As Button

        objListBoxButtonColl.Add(New Button())

        With objListBoxButtonColl(iIndex)
            .Name = "ListBoxButton" & iIndex
            .Size = New Size(iWidth, iHeight)
            .Location = New Point(iLeft, iTop)
            .Text = "Clear Selected " & cCaption
        End With

        objButton = CType(objListBoxButtonColl(iIndex), Button)

        AddHandler objButton.Click, AddressOf objListBox_Click

        Me.Controls.AddRange(New System.Windows.Forms.Control() {objListBoxButtonColl(iIndex)})

    End Sub

    Private Sub objListBox_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim objListbox As ListBox
        Dim objButton As Button
        Dim iIndex As Integer

        For Each objButton In objListBoxButtonColl

            iIndex += 1

            If objButton Is sender Then
                Exit For
            End If

        Next

        objListbox = CType(objListBoxColl(iIndex), ListBox)

        objListbox.ClearSelected()

    End Sub

    Private Sub ShowComboBox(ByVal iIndex As Short, ByVal iLeft As Short, _
        ByVal iTop As Short, ByVal iWidth As Short, ByVal iHeight As Short, ByVal cCaption As String)

        Call AddDynamicComboBoxLabel(iIndex, iLeft, iTop, iWidth, iHeight, cCaption)

        Call AddDynamicComboBox(iIndex, objComboBoxLabelColl(iIndex).Left + _
            objComboBoxLabelColl(iIndex).Width + 5, iTop, iWidth, iHeight)

    End Sub

    Private Sub AddDynamicComboBoxLabel(ByVal iIndex As Short, ByVal iLeft As Short, _
        ByVal iTop As Short, ByVal iWidth As Short, ByVal iHeight As Short, ByVal cCaption As String)

        objComboBoxLabelColl.Add(New Label())

        With objComboBoxLabelColl(iIndex)
            .AutoSize = True
            .Name = "ComboBoxLabel" & iIndex
            .Size = New Size(iWidth, iHeight)
            .Location = New Point(iLeft, iTop)
            .Text = cCaption
        End With

        Me.Controls.AddRange(New System.Windows.Forms.Control() {objComboBoxLabelColl(iIndex)})

    End Sub

    Private Sub AddDynamicComboBox(ByVal iIndex As Short, ByVal iLeft As Short, _
        ByVal iTop As Short, ByVal iWidth As Short, ByVal iHeight As Short)

        objComboBoxColl.Add(New ComboBox())

        With objComboBoxColl(iIndex)
            .Name = "ComboBox" & iIndex
            .Size = New Size(iWidth, iHeight)
            .Location = New Point(iLeft, iTop)
            .DropDownStyle = ComboBoxStyle.DropDownList
        End With

        Me.Controls.AddRange(New System.Windows.Forms.Control() {objComboBoxColl(iIndex)})

    End Sub

    Private Sub ShowDateRange(ByVal iIndex As Short, ByVal iLeft As Short, _
        ByVal iTop As Short, ByVal iHeight As Short, ByVal cCaption As String)

        Call AddDynamicDateRangeLabel(iIndex, iLeft, iTop, iHeight, cCaption)

        Call AddDynamicDateFrom(iIndex, objDateRangeLabelColl(iIndex).Left + _
            objDateRangeLabelColl(iIndex).Width + 5, iTop, iHeight)

        Call AddDynamicDateTo(iIndex, objDateFromColl(iIndex).Left + _
            objDateFromColl(iIndex).Width + 5, iTop, iHeight)

    End Sub

    Private Sub AddDynamicDateRangeLabel(ByVal iIndex As Short, ByVal iLeft As Short, _
        ByVal iTop As Short, ByVal iHeight As Short, ByVal cCaption As String)

        objDateRangeLabelColl.Add(New Label())

        With objDateRangeLabelColl(iIndex)
            .AutoSize = True
            .Name = "DateRangeLabel" & iIndex
            .Height = iHeight
            .Location = New Point(iLeft, iTop)
            .Text = cCaption
        End With

        Me.Controls.AddRange(New System.Windows.Forms.Control() {objDateRangeLabelColl(iIndex)})

    End Sub

    Private Sub AddDynamicDateFrom(ByVal iIndex As Short, ByVal iLeft As Short, _
        ByVal iTop As Short, ByVal iHeight As Short)

        objDateFromColl.Add(New DateTimePicker())

        With objDateFromColl(iIndex)
            .Name = "DateFrom" & iIndex
            .Size = New Size(90, iHeight)
            .Location = New Point(iLeft, iTop)
            .Format = DateTimePickerFormat.Custom
            .CustomFormat = "MM/dd/yyyy"
        End With

        Me.Controls.AddRange(New System.Windows.Forms.Control() {objDateFromColl(iIndex)})

    End Sub

    Private Sub AddDynamicDateTo(ByVal iIndex As Short, ByVal iLeft As Short, _
        ByVal iTop As Short, ByVal iHeight As Short)

        objDateToColl.Add(New DateTimePicker())

        With objDateToColl(iIndex)
            .Name = "DateTo" & iIndex
            .Size = New Size(90, iHeight)
            .Location = New Point(iLeft, iTop)
            .Format = DateTimePickerFormat.Custom
            .CustomFormat = "MM/dd/yyyy"
        End With

        Me.Controls.AddRange(New System.Windows.Forms.Control() {objDateToColl(iIndex)})

    End Sub

    Private Sub ShowCheckBox(ByVal iIndex As Short, ByVal iLeft As Short, ByVal iTop As Short, _
        ByVal iHeight As Short, ByVal cCaption As String)

        Call AddDynamicCheckBox(iIndex, iLeft, iTop, iHeight, cCaption)

    End Sub

    Private Sub AddDynamicCheckBox(ByVal iIndex As Short, ByVal iLeft As Short, ByVal iTop As Short, _
        ByVal iHeight As Short, ByVal cCaption As String)
        Dim objLabel As New Label()
        Dim iWidth As Short

        objCheckBoxColl.Add(New CheckBox())

        With objLabel
            .Visible = False
            .Font() = objCheckBoxColl.Item(iIndex).Font
            .Text = cCaption
            .AutoSize = True
            iWidth = .Width + 25
        End With

        objLabel = Nothing

        With objCheckBoxColl.Item(iIndex)
            .Name = "CheckBox" & iIndex
            .Size = New Size(iWidth, iHeight)
            .Location = New Point(iLeft, iTop)
            .ThreeState = True
            .CheckState = CheckState.Indeterminate
            .Text = cCaption
        End With

        Me.Controls.AddRange(New System.Windows.Forms.Control() {objCheckBoxColl(iIndex)})

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub


    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Dim objSQL As New StringBuilder()
        Dim cSQL As String
        Dim objCriteria As New StringBuilder()
        Dim cCriteria As String
        Dim cDateFrom As String
        Dim cDateTo As String
        Dim iDateCnt As Short
        Dim x As Short

        iDateCnt = objDateFromColl.Count

        For x = 1 To iDateCnt

            cDateFrom = objDateFromColl(x).text
            cDateTo = objDateToColl(x).text

            If cDateFrom <> String.Empty Or cDateTo <> String.Empty Then

                If Not IsDate(cDateFrom) Or Not IsDate(cDateTo) Then

                    MsgBox("You must enter both a valid 'from' and a 'to' date.", vbOKOnly)

                    Exit Sub

                End If

            End If


            If IsDate(cDateFrom) And IsDate(cDateTo) Then

                If DateValue(cDateFrom) > DateValue(cDateTo) Then

                    MsgBox("The 'from date' cannot be greater than the 'to date'.", vbOKOnly)

                    objDateFromColl(x).Focus()

                    Exit Sub

                End If

            End If

        Next x

        cDateFrom = objDateFromColl(SELECT_ENTER_DATE).Text
        cDateTo = objDateToColl(SELECT_ENTER_DATE).Text

        With objSQL
            .AppendFormat(GetDateSQL("createdate", cDateFrom, cDateTo))
            .AppendFormat(GetListBoxSQL(objListBoxColl(SELECT_PRODUCT), "productid"))
            .AppendFormat(GetListBoxSQL(objListBoxColl(SELECT_SOURCE), "sourceid"))
            .AppendFormat(GetComboBoxSQL(SELECT_DEPARTMENT, "deptid"))
            .AppendFormat(GetCheckBoxSQL(SELECT_COMPLEX_QUESTION, "questiontype"))
        End With

        With objCriteria
            .AppendFormat(GetDateCriteria("createdate", cDateFrom, cDateTo))
            .AppendFormat(GetListBoxCriteria(objListBoxColl(SELECT_PRODUCT), "Product"))
            .AppendFormat(GetListBoxCriteria(objListBoxColl(SELECT_SOURCE), "Source"))
            .AppendFormat(GetComboBoxCriteria(SELECT_DEPARTMENT, "Department"))
            .AppendFormat(GetCheckBoxCriteria(SELECT_COMPLEX_QUESTION, "Complex Question"))
        End With

        cSQL = objSQL.ToString
        cCriteria = objCriteria.ToString

        If objSQL.ToString <> String.Empty Then
            cSQL = Mid(cSQL, 1, cSQL.Length - 4)
        End If

        If objCriteria.ToString <> String.Empty Then
            cCriteria = Mid$(cCriteria, 1, objCriteria.Length - 5)
        End If

        MsgBox(cSQL.ToString)

        MsgBox(cCriteria.ToString)

    End Sub

    Function GetComboBoxSQL(ByVal Index As Integer, ByVal cColumn As String) As String
        Dim cSQL As String = String.Empty
        Dim lChoiceID As Long

        lChoiceID = objComboBoxColl(Index).SelectedValue()

        If lChoiceID > 0 Then

            cSQL = cColumn & " = " & lChoiceID & " AND "

        End If

        GetComboBoxSQL = cSQL

    End Function

    Function GetComboBoxCriteria(ByVal Index As Integer, ByVal cDescr As String) As String
        Dim cResult As String = String.Empty
        Dim cChoice As String

        cChoice = objComboBoxColl(Index).Text

        If cChoice <> String.Empty Then

            cResult = "the " & cDescr & " is " & cChoice & " and "

        End If

        GetComboBoxCriteria = cResult

    End Function

    Function GetCheckBoxSQL(ByVal iIndex As Integer, ByVal cColumn As String) As String
        Dim cSQL As String = String.Empty

        Select Case objCheckBoxColl(iIndex).CheckState

            Case CheckState.Checked
                cSQL = cColumn & " = " & BOOL_TRUE & " AND "

            Case CheckState.Unchecked
                cSQL = cColumn & " = " & BOOL_FALSE & " AND "

        End Select

        GetCheckBoxSQL = cSQL

    End Function

    Function GetCheckBoxCriteria(ByVal iIndex As Integer, ByVal cDescr As String) As String
        Dim cResult As String = String.Empty

        Select Case objCheckBoxColl(iIndex).CheckState

            Case CheckState.Checked
                cResult = "the " & cDescr & " is True and "

            Case CheckState.Unchecked
                cResult = "the " & cDescr & " is False and "

        End Select

        GetCheckBoxCriteria = cResult

    End Function

    Function GetDateSQL(ByVal cFieldName As String, ByVal cDateFrom As String, _
        ByVal cDateTo As String) As String
        Dim cSQL As String = String.Empty

        If IsDate(cDateFrom) And IsDate(cDateTo) Then

            cSQL = cFieldName & " BETWEEN " & Chr(39) & cDateFrom & Chr(39) & _
                " AND " & Chr(39) & cDateTo & Chr(39) & " AND "

        End If

        Return cSQL

    End Function

    Function GetDateCriteria(ByVal cDescr As String, ByVal cDateFrom As String, _
        ByVal cDateTo As String) As String
        Dim cResult As String = String.Empty

        If IsDate(cDateFrom) And IsDate(cDateTo) Then

            cResult = cResult & "the " & cDescr & " is between " & cDateFrom & _
                " and " & cDateTo & " and " & vbCrLf

        End If

        Return cResult

    End Function

    Function GetListBoxSQL(ByRef objListBox As ListBox, ByVal cColumn As String) As String
        Dim cSQL As String = String.Empty
        Dim cList As String

        If objListBox.SelectedItems.Count > 0 Then

            cList = ParseIt(objListBox, True, False, 0)

            cSQL = cColumn & " IN " & cList & " AND "

        End If

        GetListBoxSQL = cSQL

    End Function

    Function GetListBoxCriteria(ByRef objListBox As ListBox, ByVal cDescr As String) As String
        Dim cResult As String = String.Empty
        Dim cNames As String

        If objListBox.SelectedItems.Count > 0 Then

            cNames = ParseIt(objListBox, True, False, 1)

            cResult = cResult & "the " & cDescr & " is among " & cNames & " and " & vbCrLf

        End If

        GetListBoxCriteria = cResult

    End Function

    Function ParseIt(ByVal oList As ListBox, ByVal bTagged As Boolean, ByVal bQuotes As Boolean, _
        ByVal iCol As Short) As String
        Dim objResult As New StringBuilder("(")
        Dim cResult As String = String.Empty
        Dim cQuotes As String
        Dim cData As String
        Dim oTemp As Object
        Dim oCollection As Object

        If bQuotes Then
            cQuotes = Chr(39)
        Else
            cQuotes = String.Empty
        End If

        If bTagged Then
            oCollection = oList.SelectedItems
        Else
            oCollection = oList.Items
        End If

        For Each oTemp In oCollection

            cData = oTemp.Item(iCol)

            objResult.AppendFormat(cQuotes & cData & cQuotes & ",")

        Next

        cResult = objResult.ToString

        If bQuotes Then
            cResult = Mid$(cResult, 1, Len(cResult) - 2) & cQuotes & ")"
        Else
            cResult = Mid$(cResult, 1, Len(cResult)) & cQuotes & ")"
        End If

        cResult = Replace(cResult, ",)", ")")

        cResult = Replace(cResult, ",,", String.Empty)

        If cResult = "()" Then
            cResult = String.Empty
        End If

        ParseIt = cResult

    End Function

End Class


