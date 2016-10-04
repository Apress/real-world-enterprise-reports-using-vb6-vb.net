VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFillGrid 
      Caption         =   "Fill Grid"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCatalog 
      Caption         =   "Catalog"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   6840
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   11668
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objConn As New ADODB.Connection
Dim objCellset As New ADOMD.Cellset
    
Private Sub cmdCatalog_Click()
    Dim objCubeDef As ADOMD.CubeDef
    Dim objDimension As ADOMD.Dimension
    Dim objCatalog As ADOMD.Catalog
    Dim cMsg As String
    
    'Create a Catalog object
    Set objCatalog = New ADOMD.Catalog
    Set objCatalog.ActiveConnection = objConn
    
    'and set a reference to the Sales Cube
    Set objCubeDef = objCatalog.CubeDefs("Sales")
    
    'Extract the names of the Sales cube Dimensions
    For Each objDimension In objCubeDef.Dimensions
        cMsg = cMsg & objDimension.Name & vbCrLf
    Next objDimension
    
    '...and show 'em to the user
    MsgBox "Dimensions for the " & objCubeDef.Name & _
            " cube are: " & vbCrLf & cMsg

End Sub


Public Sub OpenOLAPConnection()

    ' Set up the connection to string the server.
    objConn.ConnectionString = "Datasource=LocalHost; Provider=msolap; " & _
                             "Initial Catalog=Northwind;"
    
    objConn.Open
    
    ' Create the Cellset Active Connection Object
    Set objCellset.ActiveConnection = objConn
    
    ' Create the MDX Query
    objCellset.Source = "SELECT Measures.MEMBERS ON COLUMNS," & _
                        "{[Geography].[Country].[Germany].CHILDREN," & _
                        "[Geography].[Country].[France].CHILDREN} ON ROWS " & _
                        "From [Sales]"
        
    ' Open Object
    objCellset.Open
    
End Sub

Sub LoadGrid()
    Dim iCols As Integer
    Dim iRows As Integer
    Dim iColCount As Integer
    Dim iRowCount As Integer
    Dim oPos As ADOMD.Position
    Dim x As Integer
    
    With MSFlexGrid1
    
        iColCount = objCellset.Axes(0).Positions.Count
        iRowCount = objCellset.Axes(1).Positions.Count
                
        .Cols = iColCount + 1
        .Rows = iRowCount + 1
        
        .ColWidth(0) = 1500
    
    
        'Get the column headers and align numeric data to the right
        
        x = 1
        
        For Each oPos In objCellset.Axes(0).Positions
            .TextMatrix(0, x) = oPos.Members(0).Caption
            .ColAlignment(x) = flexAlignRightCenter
            x = x + 1
        Next
    
    
        'Get the row headers.
    
        x = 0
        
        For Each oPos In objCellset.Axes(1).Positions
            .TextMatrix(x + 1, 0) = oPos.Members(0).Caption
            x = x + 1
        Next
    
        'Populate the individual dollar values
        For iCols = 0 To iColCount - 1
            For iRows = 0 To iRowCount - 1
                .TextMatrix(iRows + 1, iCols + 1) = _
                    Format(objCellset(iCols, iRows).Value, "$#0.00")
            Next iRows
        Next iCols
    
    End With
    
End Sub

Private Sub cmdFillGrid_Click()
    
    Call LoadGrid
    
End Sub

Private Sub Form_Load()

    ' Open the Connection to the OLAP Server
    Call OpenOLAPConnection
    
End Sub
