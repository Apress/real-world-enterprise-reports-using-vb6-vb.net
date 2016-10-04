VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
Dim objOutlook As Outlook.Application
Dim objNameSpace As Outlook.NameSpace
Dim objMailItem As Outlook.MailItem
Dim objAttachments As Outlook.Attachments

Dim cCC As String
Dim cSubject As String
Dim cBody As String

Public Function AddRecipient(ByVal vData As String)
    objMailItem.Recipients.Add vData
End Function

Public Property Let CC(ByVal vData As String)
    objMailItem.CC = vData
    
    cCC = vData
End Property

Public Property Get CC() As String
    CC = cCC
End Property

Public Property Let Subject(ByVal vData As String)
    objMailItem.Subject = vData
    
    cSubject = vData
End Property

Public Property Get Subject() As String
    Subject = cSubject
End Property

Public Property Let Body(ByVal vData As String)
    objMailItem.Body = vData
    
    cBody = vData
End Property

Public Property Get Body() As String
    Body = cBody
End Property

Public Function StartOutlook()

    Set objOutlook = CreateObject("Outlook.Application")
    Set objNameSpace = objOutlook.GetNamespace("MAPI")
    Set objMailItem = objOutlook.CreateItem(olMailItem)
    Set objAttachments = objMailItem.Attachments
    
End Function

Public Function CloseOutlook()
    
    objOutlook.Quit

    Set objOutlook = Nothing
    Set objNameSpace = Nothing
    Set objMailItem = Nothing
    Set objAttachments = Nothing
    
End Function

Public Function Save()
    objMailItem.Save
End Function

Public Function Send()
    objMailItem.Send
End Function

Public Function AddAttachments(cSource As String, iType As Integer, iPosition As Integer, cDisplay As String)
    Call objAttachments.Add(cSource, iType, iPosition, cDisplay)
End Function

Public Function ResolveAll()
    ResolveAll = objMailItem.Recipients.ResolveAll
End Function


