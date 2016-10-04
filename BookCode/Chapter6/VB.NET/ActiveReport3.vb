Imports System
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document

Public Class ActiveReport3
Inherits ActiveReport
	Public Sub New()
	MyBase.New()
		InitializeReport()
	End Sub
	#Region "ActiveReports Designer generated code"
	Private Sub InitializeReport()
		Me.LoadLayout(Me.GetType(),"ActiveReport3.rpx")
	End Sub
	#End Region
End Class
