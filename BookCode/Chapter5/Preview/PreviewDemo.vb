Imports C1.C1PrintDocument
Imports C1.C1PrintDocument.Util
Imports System.IO
Imports System.ComponentModel
Imports System.Xml
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Drawing.Printing

Public Enum LetterType
    ltPlainText = 1
    ltRTF = 2
End Enum

Public Class frmPreviewDemo
    Inherits System.Windows.Forms.Form

    Dim oConn As New OleDbConnection()
    Dim cConnectString As String

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
    Friend WithEvents Doc As C1.C1PrintDocument.C1PrintDocument
    Friend WithEvents C1PrintPreview1 As C1.Win.C1PrintPreview.C1PrintPreview
    Friend WithEvents PreviewToolBarButton1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton2 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton3 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton4 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton5 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton6 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton7 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton8 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton9 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton10 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton11 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton12 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton13 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton14 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton15 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton16 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton17 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton18 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton19 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton20 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton21 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton22 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton23 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton24 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton25 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton26 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton27 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton28 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton29 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton30 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton31 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton32 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents cmdSelectPrinter As System.Windows.Forms.Button
    Friend WithEvents lstPrinters As System.Windows.Forms.ListBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents Button10 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Doc = New C1.C1PrintDocument.C1PrintDocument()
        Me.C1PrintPreview1 = New C1.Win.C1PrintPreview.C1PrintPreview()
        Me.PreviewToolBarButton1 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton2 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton3 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton4 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton5 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton6 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton7 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton8 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton9 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton10 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton11 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton12 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton13 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton14 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton15 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton16 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton17 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton18 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton19 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton20 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton21 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton22 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton23 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton24 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton25 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton26 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton27 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton28 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton29 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton30 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton31 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.PreviewToolBarButton32 = New C1.Win.C1PrintPreview.PreviewToolBarButton()
        Me.cmdSelectPrinter = New System.Windows.Forms.Button()
        Me.lstPrinters = New System.Windows.Forms.ListBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Button10 = New System.Windows.Forms.Button()
        CType(Me.C1PrintPreview1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Doc
        '
        Me.Doc.C1DPageSettings = "color:True;landscape:False;margins:100,100,100,100;papersize:850,1100,TABlAHQAdAB" & _
        "lAHIAIAA4ACAAMQAvADIAIAB4ACAAMQAxACAAaQBuAA=="
        Me.Doc.ColumnSpacingStr = "0.5in"
        Me.Doc.ColumnSpacingUnit.DefaultType = True
        Me.Doc.ColumnSpacingUnit.UnitValue = "0.5in"
        Me.Doc.DefaultUnit = C1.C1PrintDocument.UnitTypeEnum.Inch
        Me.Doc.DocumentName = ""
        '
        'C1PrintPreview1
        '
        Me.C1PrintPreview1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.C1PrintPreview1.C1DPageSettings = "color:True;landscape:False;margins:100,100,100,100;papersize:850,1100,TABlAHQAdAB" & _
        "lAHIAIAA4ACAAMQAvADIAIAB4ACAAMQAxACAAaQBuAA=="
        Me.C1PrintPreview1.Document = Me.Doc
        Me.C1PrintPreview1.Location = New System.Drawing.Point(8, 8)
        Me.C1PrintPreview1.Name = "C1PrintPreview1"
        Me.C1PrintPreview1.NavigationBar.Cursor = System.Windows.Forms.Cursors.Default
        Me.C1PrintPreview1.NavigationBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.C1PrintPreview1.NavigationBar.OutlineView.Cursor = System.Windows.Forms.Cursors.Default
        Me.C1PrintPreview1.NavigationBar.OutlineView.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.C1PrintPreview1.NavigationBar.OutlineView.Indent = 19
        Me.C1PrintPreview1.NavigationBar.OutlineView.ItemHeight = 16
        Me.C1PrintPreview1.NavigationBar.OutlineView.TabIndex = 0
        Me.C1PrintPreview1.NavigationBar.OutlineView.Visible = False
        Me.C1PrintPreview1.NavigationBar.Padding = New System.Drawing.Point(6, 3)
        Me.C1PrintPreview1.NavigationBar.TabIndex = 2
        Me.C1PrintPreview1.NavigationBar.ThumbnailsView.AutoArrange = True
        Me.C1PrintPreview1.NavigationBar.ThumbnailsView.Cursor = System.Windows.Forms.Cursors.Default
        Me.C1PrintPreview1.NavigationBar.ThumbnailsView.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.C1PrintPreview1.NavigationBar.ThumbnailsView.TabIndex = 0
        Me.C1PrintPreview1.NavigationBar.ThumbnailsView.Visible = True
        Me.C1PrintPreview1.NavigationBar.Width = 160
        Me.C1PrintPreview1.PreviewPane.StartPageIdx = 0
        Me.C1PrintPreview1.PreviewPane.ZoomFactor = 0.4185906!
        Me.C1PrintPreview1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.PageArray
        Me.C1PrintPreview1.PreviewPane.ZoomPercent = 42
        Me.C1PrintPreview1.Size = New System.Drawing.Size(576, 496)
        Me.C1PrintPreview1.Splitter.Cursor = System.Windows.Forms.Cursors.VSplit
        Me.C1PrintPreview1.Splitter.Width = 3
        Me.C1PrintPreview1.StatusBar.Cursor = System.Windows.Forms.Cursors.Default
        Me.C1PrintPreview1.StatusBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.C1PrintPreview1.StatusBar.TabIndex = 4
        Me.C1PrintPreview1.StatusBar.Text = "Ready"
        Me.C1PrintPreview1.TabIndex = 3
        Me.C1PrintPreview1.ToolBar.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.PreviewToolBarButton1, Me.PreviewToolBarButton2, Me.PreviewToolBarButton3, Me.PreviewToolBarButton4, Me.PreviewToolBarButton5, Me.PreviewToolBarButton6, Me.PreviewToolBarButton7, Me.PreviewToolBarButton8, Me.PreviewToolBarButton9, Me.PreviewToolBarButton10, Me.PreviewToolBarButton11, Me.PreviewToolBarButton12, Me.PreviewToolBarButton13, Me.PreviewToolBarButton14, Me.PreviewToolBarButton15, Me.PreviewToolBarButton16, Me.PreviewToolBarButton17, Me.PreviewToolBarButton18, Me.PreviewToolBarButton19, Me.PreviewToolBarButton20, Me.PreviewToolBarButton21, Me.PreviewToolBarButton22, Me.PreviewToolBarButton23, Me.PreviewToolBarButton24, Me.PreviewToolBarButton25, Me.PreviewToolBarButton26, Me.PreviewToolBarButton27, Me.PreviewToolBarButton28, Me.PreviewToolBarButton29, Me.PreviewToolBarButton30, Me.PreviewToolBarButton31, Me.PreviewToolBarButton32})
        Me.C1PrintPreview1.ToolBar.Cursor = System.Windows.Forms.Cursors.Default
        Me.C1PrintPreview1.ToolBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.C1PrintPreview1.ToolBar.ImageSet = C1.Win.C1PrintPreview.ToolBarImageSetEnum.RegularCool
        '
        'PreviewToolBarButton1
        '
        Me.PreviewToolBarButton1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.FileOpen
        Me.PreviewToolBarButton1.ImageIndex = 0
        Me.PreviewToolBarButton1.ToolTipText = "File Open"
        '
        'PreviewToolBarButton2
        '
        Me.PreviewToolBarButton2.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.FileSave
        Me.PreviewToolBarButton2.ImageIndex = 1
        Me.PreviewToolBarButton2.ToolTipText = "File Save"
        '
        'PreviewToolBarButton3
        '
        Me.PreviewToolBarButton3.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.FilePrint
        Me.PreviewToolBarButton3.ImageIndex = 2
        Me.PreviewToolBarButton3.ToolTipText = "Print"
        '
        'PreviewToolBarButton4
        '
        Me.PreviewToolBarButton4.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.PageSetup
        Me.PreviewToolBarButton4.ImageIndex = 3
        Me.PreviewToolBarButton4.ToolTipText = "Page Setup"
        '
        'PreviewToolBarButton5
        '
        Me.PreviewToolBarButton5.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.Reflow
        Me.PreviewToolBarButton5.ImageIndex = 4
        Me.PreviewToolBarButton5.ToolTipText = "Reflow"
        '
        'PreviewToolBarButton6
        '
        Me.PreviewToolBarButton6.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.Stop
        Me.PreviewToolBarButton6.ImageIndex = 5
        Me.PreviewToolBarButton6.ToolTipText = "Stop"
        Me.PreviewToolBarButton6.Visible = False
        '
        'PreviewToolBarButton7
        '
        Me.PreviewToolBarButton7.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.PreviewToolBarButton7.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'PreviewToolBarButton8
        '
        Me.PreviewToolBarButton8.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ShowNavigationBar
        Me.PreviewToolBarButton8.ImageIndex = 6
        Me.PreviewToolBarButton8.Pushed = True
        Me.PreviewToolBarButton8.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton8.ToolTipText = "Show Navigation Bar"
        '
        'PreviewToolBarButton9
        '
        Me.PreviewToolBarButton9.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.PreviewToolBarButton9.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'PreviewToolBarButton10
        '
        Me.PreviewToolBarButton10.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.MouseHand
        Me.PreviewToolBarButton10.ImageIndex = 7
        Me.PreviewToolBarButton10.Pushed = True
        Me.PreviewToolBarButton10.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton10.ToolTipText = "Hand Tool"
        '
        'PreviewToolBarButton11
        '
        Me.PreviewToolBarButton11.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.MouseZoom
        Me.PreviewToolBarButton11.ImageIndex = 8
        Me.PreviewToolBarButton11.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton11.ToolTipText = "Zoom Tool"
        '
        'PreviewToolBarButton12
        '
        Me.PreviewToolBarButton12.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.MouseSelect
        Me.PreviewToolBarButton12.Enabled = False
        Me.PreviewToolBarButton12.ImageIndex = 9
        Me.PreviewToolBarButton12.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton12.ToolTipText = "Select Text"
        Me.PreviewToolBarButton12.Visible = False
        '
        'PreviewToolBarButton13
        '
        Me.PreviewToolBarButton13.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.FindText
        Me.PreviewToolBarButton13.ImageIndex = 10
        Me.PreviewToolBarButton13.ToolTipText = "Find Text"
        '
        'PreviewToolBarButton14
        '
        Me.PreviewToolBarButton14.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.PreviewToolBarButton14.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'PreviewToolBarButton15
        '
        Me.PreviewToolBarButton15.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.GoFirst
        Me.PreviewToolBarButton15.Enabled = False
        Me.PreviewToolBarButton15.ImageIndex = 11
        Me.PreviewToolBarButton15.ToolTipText = "First Page"
        '
        'PreviewToolBarButton16
        '
        Me.PreviewToolBarButton16.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.GoPrev
        Me.PreviewToolBarButton16.Enabled = False
        Me.PreviewToolBarButton16.ImageIndex = 12
        Me.PreviewToolBarButton16.ToolTipText = "Previous Page"
        '
        'PreviewToolBarButton17
        '
        Me.PreviewToolBarButton17.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.GoNext
        Me.PreviewToolBarButton17.ImageIndex = 13
        Me.PreviewToolBarButton17.ToolTipText = "Next Page"
        '
        'PreviewToolBarButton18
        '
        Me.PreviewToolBarButton18.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.GoLast
        Me.PreviewToolBarButton18.ImageIndex = 14
        Me.PreviewToolBarButton18.ToolTipText = "Last Page"
        '
        'PreviewToolBarButton19
        '
        Me.PreviewToolBarButton19.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.PreviewToolBarButton19.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'PreviewToolBarButton20
        '
        Me.PreviewToolBarButton20.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.HistoryPrev
        Me.PreviewToolBarButton20.Enabled = False
        Me.PreviewToolBarButton20.ImageIndex = 15
        Me.PreviewToolBarButton20.ToolTipText = "Previous View"
        Me.PreviewToolBarButton20.Visible = False
        '
        'PreviewToolBarButton21
        '
        Me.PreviewToolBarButton21.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.HistoryNext
        Me.PreviewToolBarButton21.Enabled = False
        Me.PreviewToolBarButton21.ImageIndex = 16
        Me.PreviewToolBarButton21.ToolTipText = "Next View"
        Me.PreviewToolBarButton21.Visible = False
        '
        'PreviewToolBarButton22
        '
        Me.PreviewToolBarButton22.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.PreviewToolBarButton22.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        Me.PreviewToolBarButton22.Visible = False
        '
        'PreviewToolBarButton23
        '
        Me.PreviewToolBarButton23.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ZoomOut
        Me.PreviewToolBarButton23.ImageIndex = 17
        Me.PreviewToolBarButton23.ToolTipText = "Zoom Out"
        Me.PreviewToolBarButton23.Visible = False
        '
        'PreviewToolBarButton24
        '
        Me.PreviewToolBarButton24.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ZoomIn
        Me.PreviewToolBarButton24.ImageIndex = 18
        Me.PreviewToolBarButton24.ToolTipText = "Zoom In"
        Me.PreviewToolBarButton24.Visible = False
        '
        'PreviewToolBarButton25
        '
        Me.PreviewToolBarButton25.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.PreviewToolBarButton25.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        Me.PreviewToolBarButton25.Visible = False
        '
        'PreviewToolBarButton26
        '
        Me.PreviewToolBarButton26.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewActualSize
        Me.PreviewToolBarButton26.ImageIndex = 19
        Me.PreviewToolBarButton26.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton26.ToolTipText = "Actual Size"
        '
        'PreviewToolBarButton27
        '
        Me.PreviewToolBarButton27.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewFullPage
        Me.PreviewToolBarButton27.ImageIndex = 20
        Me.PreviewToolBarButton27.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton27.ToolTipText = "Full Page"
        '
        'PreviewToolBarButton28
        '
        Me.PreviewToolBarButton28.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewPageWidth
        Me.PreviewToolBarButton28.ImageIndex = 21
        Me.PreviewToolBarButton28.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton28.ToolTipText = "Page Width"
        '
        'PreviewToolBarButton29
        '
        Me.PreviewToolBarButton29.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewTwoPages
        Me.PreviewToolBarButton29.ImageIndex = 22
        Me.PreviewToolBarButton29.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton29.ToolTipText = "Two Pages"
        '
        'PreviewToolBarButton30
        '
        Me.PreviewToolBarButton30.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewFourPages
        Me.PreviewToolBarButton30.ImageIndex = 23
        Me.PreviewToolBarButton30.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.PreviewToolBarButton30.ToolTipText = "Four Pages"
        '
        'PreviewToolBarButton31
        '
        Me.PreviewToolBarButton31.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.PreviewToolBarButton31.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        Me.PreviewToolBarButton31.Visible = False
        '
        'PreviewToolBarButton32
        '
        Me.PreviewToolBarButton32.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.Help
        Me.PreviewToolBarButton32.ImageIndex = 24
        Me.PreviewToolBarButton32.ToolTipText = "Help"
        Me.PreviewToolBarButton32.Visible = False
        '
        'cmdSelectPrinter
        '
        Me.cmdSelectPrinter.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdSelectPrinter.Location = New System.Drawing.Point(592, 480)
        Me.cmdSelectPrinter.Name = "cmdSelectPrinter"
        Me.cmdSelectPrinter.Size = New System.Drawing.Size(136, 24)
        Me.cmdSelectPrinter.TabIndex = 6
        Me.cmdSelectPrinter.Text = "Select Printer"
        '
        'lstPrinters
        '
        Me.lstPrinters.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.lstPrinters.Location = New System.Drawing.Point(592, 392)
        Me.lstPrinters.Name = "lstPrinters"
        Me.lstPrinters.Size = New System.Drawing.Size(136, 82)
        Me.lstPrinters.TabIndex = 5
        '
        'Button1
        '
        Me.Button1.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button1.Location = New System.Drawing.Point(592, 16)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(136, 24)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Getting Started"
        '
        'Button2
        '
        Me.Button2.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button2.Location = New System.Drawing.Point(592, 48)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(136, 24)
        Me.Button2.TabIndex = 8
        Me.Button2.Text = "Simple Columnar Report"
        '
        'Button3
        '
        Me.Button3.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button3.Location = New System.Drawing.Point(592, 80)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(136, 24)
        Me.Button3.TabIndex = 9
        Me.Button3.Text = "Get Printer List"
        '
        'Button4
        '
        Me.Button4.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button4.Location = New System.Drawing.Point(592, 112)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(136, 24)
        Me.Button4.TabIndex = 10
        Me.Button4.Text = "Labels"
        '
        'Button5
        '
        Me.Button5.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button5.Location = New System.Drawing.Point(592, 144)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(136, 24)
        Me.Button5.TabIndex = 11
        Me.Button5.Text = "Free Form Text"
        '
        'Button7
        '
        Me.Button7.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button7.Location = New System.Drawing.Point(592, 176)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(136, 24)
        Me.Button7.TabIndex = 13
        Me.Button7.Text = "Render Text Demo"
        '
        'Button8
        '
        Me.Button8.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button8.Location = New System.Drawing.Point(592, 208)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(136, 24)
        Me.Button8.TabIndex = 14
        Me.Button8.Text = "Border Demo"
        '
        'Button9
        '
        Me.Button9.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button9.Location = New System.Drawing.Point(592, 240)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(136, 24)
        Me.Button9.TabIndex = 15
        Me.Button9.Text = "Rendering Layers"
        '
        'Button10
        '
        Me.Button10.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button10.Location = New System.Drawing.Point(592, 272)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(136, 24)
        Me.Button10.TabIndex = 16
        Me.Button10.Text = "Graphics"
        '
        'frmPreviewDemo
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(736, 509)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button10, Me.Button9, Me.Button8, Me.Button7, Me.Button5, Me.Button4, Me.Button3, Me.Button2, Me.Button1, Me.cmdSelectPrinter, Me.lstPrinters, Me.C1PrintPreview1})
        Me.Name = "frmPreviewDemo"
        Me.Text = "Form1"
        CType(Me.C1PrintPreview1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdColumnarReportDemo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ColumnarReportDemo()
    End Sub

    Private Sub ColumnarReportDemo()
        Dim emplTb As DataTable
        Dim row As DataRow
        Dim objTable As New RenderTable(Doc)
        Dim objHeaderTable As New RenderTable(Doc)
        Dim iRow As Integer
        Dim path$ = Application.StartupPath
        Dim Dataset11 As New DataSet()
        Dim tb As DataTable

        Const fileName$ = "empl.xml"

        path = path.Substring(0, path.LastIndexOf("\bin")) + "\" + fileName
        Dataset11.ReadXml(path)

        tb = Dataset11.Tables.Item("Employees")

        'Return a Data set from the sample XML employee file
        emplTb = Dataset11.Tables.Item("Employees")

        'Define the template for the table used to display the report data
        With objTable
            'Turn off any horizontal or vertical borders
            .StyleTableCell.BorderTableHorz.Empty = True
            .StyleTableCell.BorderTableVert.Empty = True
            .Style.Borders.AllEmpty = True

            'Set amount of white space to be left around column
            .Style.Spacing.Top = 0.2
            .Style.Spacing.Left = 0.1

            With .Columns
                'Add three columns to the table
                .AddSome(3)

                'Define the style of the text in the first column
                With .Item(0).StyleTableCell
                    .Font = New Font("Arial", 10, FontStyle.Bold)
                End With

                'Define the style of the text in the second column
                With .Item(1).StyleTableCell
                    .Font = New Font("Arial", 10)
                End With

                'Set the widths of all three columns
                .Item(0).Width = 1.7
                .Item(1).Width = 3
                .Item(2).Width = 1.5

            End With

        End With


        'Define a table for the header of the report
        With objHeaderTable
            .Columns.AddSome(2)
            .Columns(0).Width = Doc.PrintableAreaSize.Width / 2
            .Columns(1).Width = Doc.PrintableAreaSize.Width / 2

            .StyleTableCell.BorderTableHorz.Empty = True
            .StyleTableCell.BorderTableVert.Empty = True
            .Style.Borders.AllEmpty = True

            'Add a row
            .Body.Rows.Add()

            'Put the page x of y commands in the first column
            .Body.Cell(0, 0).RenderText.Text = "Page [@@PageNo@@] of [@@PageCount@@]"

            'Define the style for the second column and set the text
            With .Body.Cell(0, 1).RenderText
                .Style.TextAlignHorz = AlignHorzEnum.Right
                .Style.Font = New Font("Arial", 12, FontStyle.Italic)
                .Text = "Employee List"
            End With

        End With

        'Assign the table object just create to the PageHeader. 
        'Footers can be created with the PagerFooter property
        Doc.PageHeader.RenderObject = objHeaderTable

        With Doc
            .StartDoc()

            For Each row In emplTb.Select()

                'Using the table object defined previously
                With objTable.Body
                    'Add a blank row
                    .Rows.Add()

                    'Fill each of the three cells with data from the XML data set
                    .Cell(iRow, 0).RenderText.Text = row.Item("FirstName") + " " + _
                                   row.Item("LastName")

                    .Cell(iRow, 1).RenderText.Text = row.Item("Address") + ", " + _
                                   row.Item("City") + ", " + _
                                   row.Item("Country")

                    .Cell(iRow, 2).RenderText.Text = row.Item("HomePhone")
                End With

                .RenderBlock(objTable)

                iRow += 1

            Next

            .EndDoc()

        End With
    End Sub

    Private Sub StyleObjectDemo()
        Dim objStyleName As New C1DocStyle(Doc)
        Dim objStyleQuotation As New C1DocStyle(Doc)

        With objStyleName
            .Font = New Font("Arial", 20, FontStyle.Bold)
            .TextColor = Color.Black
            .BackColor = Color.Red
            With .Borders
                .All = New LineDef(Color.Black, 5)
            End With
            .TextAlignHorz = AlignHorzEnum.Center
        End With

        With objStyleQuotation
            .Font = New Font("Arial", 10, FontStyle.Italic)
            .TextColor = Color.Black
            With .Spacing
                .TopUnit.Value = 0.2
                .BottomUnit.Value = 0.1
            End With
        End With

        With Doc
            .StartDoc()

            .RenderBlockText("Abraham Lincoln", objStyleName)
            .RenderBlockText("With malice toward none, " & _
            "with charity toward all...", objStyleQuotation)

            .RenderBlockText("Bill Clinton", objStyleName)
            .RenderBlockText("I swear I thought she " & _
            "was over eighteen.", objStyleQuotation)

            .EndDoc()
        End With
    End Sub


    Private Sub cmdStyleObjectDemo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        StyleObjectDemo()
    End Sub

    Private Sub frmPreviewDemo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        cConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BookCode\SampleDatabase.mdb"

        oConn.ConnectionString = cConnectString
        oConn.Open()

    End Sub

    Private Sub GettingStarted(ByVal oConn As OleDbConnection)
        Dim objDR As OleDbDataReader
        Dim objCommand As OleDbCommand
        Dim objTable As New RenderTable(Doc)
        Dim iRow As Integer
        Dim cSQL As String

        cSQL = "SELECT * " & _
                "FROM Source " & _
                "ORDER BY descr"

        objCommand = New OleDbCommand(cSQL, oConn)
        objDR = objCommand.ExecuteReader()

        With objTable.Columns
            .AddSome(2)
            .Item(0).Width = 1.7
            .Item(1).Width = 3
        End With


        With Doc
            .StartDoc()

            While objDR.Read

                With objTable.Body
                    .Rows.Add()
                    .Cell(iRow, 0).RenderText.Text = objDR.Item("ID").ToString
                    .Cell(iRow, 1).RenderText.Text = objDR.Item("Descr").ToString
                End With

                iRow += 1

            End While

            .RenderBlock(objTable)

            .EndDoc()

        End With

        objDR.Close()

    End Sub

    Private Sub ColumnarRpt(ByVal oConn As OleDbConnection)
        Dim objDR As OleDbDataReader
        Dim objCommand As OleDbCommand
        Dim objTable As New RenderTable(Doc)
        Dim objHeader As New RenderTable(Doc)
        Dim objFooter As New RenderTable(Doc)
        Dim iRow As Integer
        Dim cSQL As String

        cSQL = "SELECT * " & _
                "FROM Source " & _
                "ORDER BY descr"

        objCommand = New OleDbCommand(cSQL, oConn)
        objDR = objCommand.ExecuteReader()

        With objHeader
            .Columns.AddSome(3)
            .Columns(0).Width = 2.5
            .Columns(1).Width = 2
            .Columns(2).Width = 1

            .Style.Borders.AllEmpty = True
            .StyleTableCell.BorderTableHorz.Empty = True
            .StyleTableCell.BorderTableVert.Empty = True

            .Body.Rows.Add()

            With .Body.Cell(0, 0).RenderText
                .Style.TextAlignHorz = AlignHorzEnum.Left
                .Style.Font = New Font("Arial", 10, FontStyle.Bold)
                .Text = "Seton Software Development, Inc."
            End With

            With .Body.Cell(0, 1).RenderText
                .Style.TextAlignHorz = AlignHorzEnum.Center
                .Style.Font = New Font("Arial", 12, FontStyle.Bold)
                .Text = "Sample Report"
            End With

            With .Body.Cell(0, 2).RenderText
                .Style.TextAlignHorz = AlignHorzEnum.Right
                .Style.Font = New Font("Arial", 10, FontStyle.Bold)
                .Text = "Page [@@PageNo@@] of [@@PageCount@@]"
            End With

        End With

        With objFooter
            .Columns.AddSome(1)
            .Columns(0).Width = Doc.PrintableAreaSize.Width

            .Style.Borders.AllEmpty = True

            .Body.Rows.Add()

            .Body.Cell(0, 0).Style.TextAlignHorz = AlignHorzEnum.Center
            .Body.Cell(0, 0).RenderText.Text = "For internal use only"

        End With

        With objTable.Columns
            .AddSome(2)
            .Item(0).Width = 1.7
            .Item(1).Width = 3
        End With

        With Doc
            .PageHeader.RenderObject = objHeader
            .PageFooter.RenderObject = objFooter


            With .PageSettings
                .Landscape = True
                .Margins.Top = 300
                .Margins.Bottom = 200
                .Margins.Right = 200
                .Margins.Left = 300
            End With

            .StartDoc()

            While objDR.Read

                With objTable.Body

                    .Rows.Add()

                    If iRow Mod 2 = 0 Then
                        .Cell(iRow, 0).Style.BackColor = Color.Red
                    End If

                    .Cell(iRow, 0).RenderText.Text = objDR.Item("ID").ToString
                    .Cell(iRow, 1).RenderText.Text = objDR.Item("Descr").ToString
                End With

                iRow += 1

            End While

            .RenderBlock(objTable)

            .EndDoc()

        End With

        objDR.Close()

    End Sub

    Sub GetPrinterList()
        Dim cPrinter As String

        For Each cPrinter In PrinterSettings.InstalledPrinters
            lstPrinters.Items.Add(cPrinter)
        Next

    End Sub

    Private Sub cmdSelectPrinter_Click(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles cmdSelectPrinter.Click

        Doc.PageSettings.PrinterSettings.PrinterName = lstPrinters.Text

    End Sub

    Sub Labels(ByVal oConn As OleDbConnection)
        Dim objLabel As New C1.C1PrintDocument.RenderTable(Doc)
        Dim objDR As OleDb.OleDbDataReader
        Dim objCommand As New OleDb.OleDbCommand()
        Dim cSQL As String
        Dim cCSZ As String
        Dim iLabel As Integer
        Dim iCol As Integer
        Dim dblLeft As Double

        cSQL = "SELECT *FROM Requester "

        With objCommand
            .Connection = oConn
            .CommandText = cSQL
            .CommandType = CommandType.Text
            .CommandTimeout = 60
            objDR = .ExecuteReader()
        End With

        iLabel = 1
        iCol = 1
        dblLeft = 0

        With Doc

            .StartDoc()

            While objDR.Read

                objLabel = New C1.C1PrintDocument.RenderTable(Doc)

                With objLabel
                    .Columns.AddSome(1)
                    .Columns(0).Width = 3.5
                    .Style.Borders.AllEmpty = True
                    .StyleTableCell.BorderTableHorz.Empty = True
                    .StyleTableCell.BorderTableVert.Empty = True
                End With

                With objLabel.Body
                    .Rows.AddSome(4)
                    .Cell(0, 0).RenderText.Text = objDR.Item("FirstName").ToString & ", " & objDR.Item("LastName").ToString
                    .Cell(1, 0).RenderText.Text = objDR.Item("Address1").ToString
                    cCSZ = objDR.Item("City").ToString & ", " & objDR.Item("State").ToString & " " & objDR.Item("Zip").ToString

                    If IsDBNull(objDR.Item("Address2")) Then
                        .Cell(2, 0).RenderText.Text = cCSZ
                    Else

                        .Cell(2, 0).RenderText.Text = objDR.Item("Address2").ToString
                        .Cell(3, 0).RenderText.Text = cCSZ
                    End If

                End With

                .RenderDirect(dblLeft, ((iLabel - 1) * 1.9) + 0.3, objLabel)

                iLabel = iLabel + 1

                If iLabel > 5 Then

                    iLabel = 1

                    Select Case iCol
                        Case 1
                            dblLeft = 4
                            iCol = iCol + 1
                        Case 2
                            dblLeft = 0
                            iCol = 1
                            .NewPage()
                    End Select
                End If

            End While

            .EndDoc()
        End With
    End Sub

    Sub FreeFormText(ByVal oConn As OleDbConnection, ByVal iLetterType As Short)
        Dim objDR As OleDbDataReader
        Dim objCommand As OleDbCommand
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

        objCommand = New OleDbCommand(cSQL, oConn)
        objDR = objCommand.ExecuteReader()


        If objDR.Read Then
            cLetterText = objDR("LetterText")
        Else
            Exit Sub
        End If

        objDR.Close()

        cSQL = "SELECT ID, Salutation, LastName, FirstName, " & _
                "Address1, Address2, City, State, Zip " & _
                "FROM Requester " & _
                "WHERE Address1 IS NOT NULL " & _
                "ORDER BY LastName"

        objCommand = New OleDbCommand(cSQL, oConn)
        objDR = objCommand.ExecuteReader()

        With Doc

            .Style.Font = New Font("Arial", 12)
            .Style.TextAlignHorz = AlignHorzEnum.Justify
            .Style.LineSpacing = 120

            .StartDoc()

            While objDR.Read

                cText = cLetterText

                cCaseNumber = objDR("id").ToString
                cSalutation = "" & objDR("Salutation").ToString
                cLastName = "" & objDR("lastname").ToString
                cName = objDR("firstname").ToString & " " & objDR("lastname").ToString
                cAddress1 = "" & objDR("Address1").ToString
                cAddress2 = "" & objDR("Address2").ToString

                If cAddress2 <> vbNullString Then
                    cAddress1 = cAddress1 & vbCrLf & cAddress2
                End If

                cCSZ = objDR("city").ToString & ", " & objDR("state").ToString & " " & objDR("zip").ToString

                cText = Replace(cText, "%Name%", cName)
                cText = Replace(cText, "%Address%", cAddress1)
                cText = Replace(cText, "%CSZ%", cCSZ)
                cText = Replace(cText, "%Salutation%", cSalutation)
                cText = Replace(cText, "%LastName%", cLastName)
                cText = Replace(cText, "%CaseNumber%", cCaseNumber)

                .RenderInlineText(cText)

                .NewPage()

            End While

            .EndDoc()

        End With

        objDR.Close()
        objDR = Nothing

    End Sub

    Sub RenderTextDemo()
        Dim objStyle As New C1DocStyle(Doc)

        With objStyle
            .Font = New Font("Arial", 20, FontStyle.Bold)
        End With

        With Doc
            .StartDoc()

            .RenderInlineText("One thing that's really cool about this in-line text feature ")
            .RenderInlineText("is the ability to change fonts ", New Font("Ariel", 24, FontStyle.Bold))
            .RenderInlineText("and text color", New Font("Ariel", 24, FontStyle.Regular), Color.Green)
            .RenderInlineEnd()
            .RenderInlineText("One thing that's really cool about this in-line text feature ")
            .RenderInlineText("is the ability to change fonts ", New Font("Ariel", 24, FontStyle.Bold))
            .RenderInlineText("and text color", New Font("Ariel", 24, FontStyle.Regular), Color.Green)

            .RenderDirectText(1, 4, "RenderDirectText is another cool feature", 20, 40, objStyle)
            .RenderDirectText(1.5, 4.5, "that prints text exactly where you tell it to", 20, 40, objStyle)

            .RenderBlockText("RenderBlockText is another great method to print chunks of string data")

            .EndDoc()

        End With

    End Sub

    Sub BorderDemo()
        Dim objTable As New RenderTable(Doc)
        Dim objHeader As New RenderTable(Doc)
        Dim objFooter As New RenderTable(Doc)
        Dim objTableBottomLineDef As New LineDef(Color.Red, 10)
        Dim objTableLeftLineDef As New LineDef(Color.Green, 10)
        Dim objTableRightLineDef As New LineDef(Color.RoyalBlue, 10)
        Dim objTableTopLineDef As New LineDef(Color.Gold, 10)
        Dim objCellBottomLineDef As New LineDef(Color.Silver, 3)
        Dim objCellLeftLineDef As New LineDef(Color.SeaShell, 3)
        Dim objCellRightLineDef As New LineDef(Color.SeaGreen, 3)
        Dim objCellTopLineDef As New LineDef(Color.Salmon, 3)
        Dim x As Short

        With objHeader
            .Columns.AddSome(3)
            .Columns(0).Width = 2
            .Columns(1).Width = 2
            .Columns(2).Width = 2

            .Style.Borders.AllEmpty = False
            .StyleTableCell.BorderTableHorz.Empty = True
            .StyleTableCell.BorderTableVert.Empty = True

            .Body.Rows.Add()

            .Body.Cell(0, 0).RenderText.Style.TextAlignHorz = AlignHorzEnum.Center
            .Body.Cell(0, 1).RenderText.Style.TextAlignHorz = AlignHorzEnum.Center
            .Body.Cell(0, 2).RenderText.Style.TextAlignHorz = AlignHorzEnum.Center

            .Body.Cell(0, 0).RenderText.Text() = "Header 1"
            .Body.Cell(0, 1).RenderText.Text() = "Header 2"
            .Body.Cell(0, 2).RenderText.Text() = "Header 3"

        End With

        With objFooter
            .Columns.AddSome(3)
            .Columns(0).Width = 2
            .Columns(1).Width = 2
            .Columns(2).Width = 2

            .Style.Borders.AllEmpty = True
            .StyleTableCell.BorderTableHorz.Empty = True
            .StyleTableCell.BorderTableVert.Empty = True

            .Body.Rows.Add()

            .Body.Cell(0, 0).RenderText.Style.TextAlignHorz = AlignHorzEnum.Center
            .Body.Cell(0, 1).RenderText.Style.TextAlignHorz = AlignHorzEnum.Center
            .Body.Cell(0, 2).RenderText.Style.TextAlignHorz = AlignHorzEnum.Center

            .Body.Cell(0, 0).RenderText.Text() = "Footer 1"
            .Body.Cell(0, 1).RenderText.Text() = "Footer 2"
            .Body.Cell(0, 2).RenderText.Text() = "Footer 3"

        End With

        objTable.Style.Borders.AllEmpty = False

        With objTable.StyleTableCell.BorderTableHorz
            .Empty = False
            .Color = Color.Indigo
            .WidthPt = 20
        End With

        With objTable.StyleTableCell.BorderTableVert
            .Empty = False
            .Color = Color.DarkGreen
            .WidthPt = 20
        End With

        With objTable.StyleTableCell.Borders
            .Bottom = objCellBottomLineDef
            .Left = objCellLeftLineDef
            .Right = objCellRightLineDef
            .Top = objCellTopLineDef
        End With

        With objTable.Style.Borders
            .Bottom = objTableBottomLineDef
            .Left = objTableLeftLineDef
            .Right = objTableRightLineDef
            .Top = objTableTopLineDef
        End With

        With objTable.Columns
            .AddSome(4)
            .Item(0).Width = 1.5
            .Item(1).Width = 1.5
            .Item(2).Width = 1.5
            .Item(3).Width = 1.5
        End With

        With Doc

            .PageHeader.RenderObject = objHeader
            .PageFooter.RenderObject = objFooter

            .StartDoc()

            For x = 0 To 10

                With objTable.Body
                    .Rows.Add()
                    .Cell(x, 0).RenderText.Text = "Col 0 - Row " & x
                    .Cell(x, 1).RenderText.Text = "Col 1 - Row " & x
                    .Cell(x, 2).RenderText.Text = "Col 2 - Row " & x
                    .Cell(x, 3).RenderText.Text = "Col 3 - Row " & x
                End With

            Next x            .RenderBlock(objTable)            .EndDoc()        End With    End Sub
    Private Sub RenderingLayers()
        Dim objCompanyLogo As New C1DocStyle(Doc)
        Dim dblHeight As Double
        Dim dblWidth As Double
        Dim cPath As String

        cPath = "c:\bookcode\chapter5\"

        With objCompanyLogo

            .BackgroundImage = Image.FromFile(cPath & "setonsoftware.jpg")

            With .BackgroundImageAlign
                .StretchHorz = False
                .StretchVert = False
                .TileHorz = True
                .TileVert = True
            End With

        End With


        With Doc
            .PageHeader.Height = 0

            .StartDoc()

            dblWidth = .BodyAreaSize.Width
            dblHeight = .BodyAreaSize.Height

            .PageLayer = DocumentPageLayerEnum.Overlay
            .RenderDirectImage(dblWidth / 10, dblHeight / 10, _
                Image.FromFile(cPath & "paidinfull.jpg"))

            .PageLayer = DocumentPageLayerEnum.Main

            .Style.TextColor = Color.Red

            .RenderBlockText("Invoice for consulting services rendered")
            .RenderBlockText("Project Management:")
            .RenderBlockText("Development:")
            .RenderBlockText("Documentation:")
            .RenderInlineEnd()

            .PageLayer = DocumentPageLayerEnum.Background
            .RenderDirectText(0, 0, " ", dblWidth, dblHeight, objCompanyLogo)

            .EndDoc()
        End With
    End Sub    Private Sub Graphics()
        Dim cPath As String
        Dim myIntArray() As Integer = {1, 2, 3, 4}
        Dim objStyle As New C1DocStyle(Doc)
        Dim objPoint As New UnitPoint()

        cPath = "c:\bookcode\chapter5\"

        objStyle.ShapeFillColor = Color.Red

        objPoint.X = 1
        objPoint.Y = 3

        With Doc

            .StartDoc()

            .RenderDirectImage(1, 7, Image.FromFile(cPath & "setonsoftware.jpg"))

            .EndDoc()

        End With

    End Sub    Sub DumpTable(ByVal oTable As RenderTable)
        Dim x As Integer
        Dim y As Integer
        Dim iRows As Integer
        Dim iCols As Integer
        Dim cLine As String
        Dim cPath As String
        Dim cFileName As String
        Dim cData As Object

        cPath = "c:\bookcode\chapter5\"

        cFileName = cPath & "dump.txt"

        FileOpen(1, cFileName, OpenMode.Output)

        With oTable

            iRows = oTable.Body.Rows.Count - 1
            iCols = oTable.Columns.Count - 1

            For x = 0 To iRows

                cLine = vbNullString

                For y = 0 To iCols

                    cData = oTable.Body.Cell(x, y).RenderText.Text

                    cLine = cLine & Chr(34) & cData & Chr(34) & ","

                Next y

                PrintLine(1, Mid$(cLine, 1, Len(cLine) - 1))

            Next x

        End With

        FileClose(1)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        GettingStarted(oConn)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ColumnarRpt(oConn)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        GetPrinterList()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Labels(oConn)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        FreeFormText(oConn, LetterType.ltPlainText)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        RenderTextDemo()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        BorderDemo()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        RenderingLayers()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Graphics()
    End Sub
End Class
