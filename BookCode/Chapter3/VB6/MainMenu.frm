VERSION 5.00
Begin VB.MDIForm frmMainMenu 
   BackColor       =   &H8000000C&
   Caption         =   "Report Demo"
   ClientHeight    =   6990
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuMain 
      Caption         =   "Reports"
      Index           =   0
      Begin VB.Menu mnuReport 
         Caption         =   "Duration Report By Product"
         Index           =   1
      End
      Begin VB.Menu mnuReport 
         Caption         =   "Duration Report By Source"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const DURATION_EDIT_BY_PRODUCT = 1
Const DURATION_EDIT_BY_SOURCE = 2

Private Sub MDIForm_Load()

    With oConn
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\BookCode\SampleDatabase.mdb;Persist Security Info=False"
        .Open
    End With

End Sub



Private Sub mnuReport_Click(Index As Integer)

    objReport.ReportTitle = mnuReport.Item(Index).Caption
    
    Select Case Index
            
        Case DURATION_EDIT_BY_PRODUCT
            objReport.HelpContext = 102
            objReport.Report = "DurationByProductRpt"
            
        Case DURATION_EDIT_BY_SOURCE
            objReport.HelpContext = 103
            objReport.Report = "DurationBySourceRpt"
        
    End Select
    
    frmReportCriteria.Show

End Sub
