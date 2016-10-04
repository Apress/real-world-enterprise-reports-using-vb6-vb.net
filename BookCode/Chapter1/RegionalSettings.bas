Attribute VB_Name = "Module1"
Option Explicit

Public Enum Locale
    LanguageID = &H1
    Language = &H2
    LanguageEnglish = &H1001
    AbbrevLanguage = &H3
    NativeLanguage = &H4
    CountryID = &H5
    Country = &H6
    CountryEnglish = &H1002
    AbbrevCountry = &H7
    NativeCountry = &H8
    DefaultLanguage = &H9
    DefaultCountryID = &HA
    DefaultCodePage = &HB
    ListItemSep = &HC
    Measure = &HD
    DecimalSep = &HE
    ThousandSep = &HF
    DigitGrouping = &H10
    FractionalDigits = &H11
    LeadingZeros = &H12
    NativeDigits = &H13
    CurrencySymbol = &H14
    IntlCurrencySymbol = &H15
    CurrDecimalSep = &H16
    CurrThousandSep = &H17
    CurrGrouping = &H18
    CurrDigits = &H19
    IntlCurrDigits = &H1A
    CurrencyPositive = &H1B
    CurrencyNegative = &H1C
    DateSep = &H1D
    TimeSep = &H1E
    ShortDateFormat = &H1F
    LongDateFormat = &H20
    TimeFormat = &H1003
    ShortDateOrdering = &H21
    LongDateOrdering = &H22
    TimeFormatSpecifier = &H23
    CenturyFormatSpecifier = &H24
    LeadingZerosTime = &H25
    LeadingZerosDate = &H26
    LeadingZerosMonth = &H27
    AMdesignator = &H28
    PMdesignator = &H29
    DayName1 = &H2A
    DayName2 = &H2B
    DayName3 = &H2C
    DayName4 = &H2D
    DayName5 = &H2E
    DayName6 = &H2F
    DayName7 = &H30
    AbbrevDayName1 = &H31
    AbbrevDayName2 = &H32
    AbbrevDayName3 = &H33
    AbbrevDayName4 = &H34
    AbbrevDayName5 = &H35
    AbbrevDayName6 = &H36
    AbbrevDayName7 = &H37
    MonthName1 = &H38
    MonthName2 = &H39
    MonthName3 = &H3A
    MonthName4 = &H3B
    MonthName5 = &H3C
    MonthName6 = &H3D
    MonthName7 = &H3E
    MonthName8 = &H3F
    MonthName9 = &H40
    MonthName10 = &H41
    MonthName11 = &H42
    MonthName12 = &H43
    AbbrevMonthName1 = &H44
    AbbrevMonthName2 = &H45
    AbbrevMonthName3 = &H46
    AbbrevMonthName4 = &H47
    AbbrevMonthName5 = &H48
    AbbrevMonthName6 = &H49
    AbbrevMonthName7 = &H4A
    AbbrevMonthName8 = &H4B
    AbbrevMonthName9 = &H4C
    AbbrevMonthName10 = &H4D
    AbbrevMonthName11 = &H4E
    AbbrevMonthName12 = &H4F
End Enum

Public Const LOCALE_SYSTEM_DEFAULT& = &H800
Public Const LOCALE_USER_DEFAULT& = &H400


Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
(ByVal Locale As Long, ByVal LCType As Long, _
ByVal lpLCData As String, ByVal cchData As Long) As Long




