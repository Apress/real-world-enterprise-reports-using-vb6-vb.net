Imports System
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document

Public Class ActiveReport2
Inherits ActiveReport
	Public Sub New()
	MyBase.New()
		InitializeReport()
	End Sub
	#Region "ActiveReports Designer generated code"
	Private Sub InitializeReport()
		Me.LoadLayout(Me.GetType(),"ActiveReport2.rpx")
	End Sub
	#End Region
End Class
