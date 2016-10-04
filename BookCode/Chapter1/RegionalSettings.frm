VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim cSettings As String

cSettings = "Abbrev. Country: " & _
    LocaleInfo(Locale.AbbrevCountry) & vbCrLf
    
cSettings = cSettings & "Country: " & _
    LocaleInfo(Locale.CountryEnglish) & vbCrLf
    
cSettings = cSettings & "Currency Decimal Separator: " & _
    LocaleInfo(Locale.CurrDecimalSep) & vbCrLf
    
cSettings = cSettings & "Language: " & _
    LocaleInfo(Locale.Language) & vbCrLf
    
cSettings = cSettings & "Currency Symbol: " & _
    LocaleInfo(Locale.IntlCurrencySymbol) & vbCrLf
    
cSettings = cSettings & "Long Date: " & _
    LocaleInfo(Locale.LongDateFormat) & vbCrLf
    
cSettings = cSettings & "Short Date: " & _
    LocaleInfo(Locale.ShortDateFormat) & vbCrLf
    
cSettings = cSettings & "Time Format: " & _
    LocaleInfo(Locale.TimeFormat) & vbCrLf

cSettings = cSettings & "First Day of week: " & _
    LocaleInfo(Locale.DayName1)
    
MsgBox cSettings
End Sub

Function LocaleInfo(lDataNeeded As Long) As String
    Dim cBuffer As String
    Dim lResult As Long
    Dim iBuffer As Integer
    
    iBuffer = 255

    cBuffer = String$(iBuffer - 1, 0)
        
    lResult = GetLocaleInfo(LOCALE_USER_DEFAULT, lDataNeeded, _
        cBuffer, iBuffer)
        
    If lResult <> 0 Then
        LocaleInfo = Left$(cBuffer, lResult - 1)
    End If
    
End Function


