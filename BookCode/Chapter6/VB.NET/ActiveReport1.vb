Imports System
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document

Public Class ActiveReport1
Inherits ActiveReport
	Public Sub New()
	MyBase.New()
		InitializeReport()
	End Sub
	#Region "ActiveReports Designer generated code"
	Private Sub InitializeReport()
		Me.LoadLayout(Me.GetType(),"ActiveReport1.rpx")
	End Sub
	#End Region
End Class
