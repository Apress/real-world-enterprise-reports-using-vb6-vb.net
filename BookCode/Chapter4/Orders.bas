Attribute VB_Name = "Module1"
Option Explicit

Function CreateReport(ByVal ReportName As String) As CRAXDRT.Report
    Dim objApplication As CRAXDRT.Application
    Dim objReport As CRAXDRT.Report
    Dim objSpecialField As CRAXDRT.FieldObject
    Dim objTextObject As CRAXDRT.TextObject
    Dim objFieldObject As CRAXDRT.FieldObject
    Dim objDBTables As CRAXDRT.DatabaseTables
    Dim objOneDBTable As CRAXDRT.DatabaseTable
    Dim objManyDBTable As CRAXDRT.DatabaseTable
    Dim objOneDBField As CRAXDRT.DatabaseFieldDefinition
    Dim objManyDBField As CRAXDRT.DatabaseFieldDefinition
    Dim objDBlinks As CRAXDRT.TableLinks
    Dim objDBLink As CRAXDRT.TableLink
    Dim objFont As Object
            
    Set objApplication = New CRAXDRT.Application
    Set objReport = objApplication.NewReport
    
    ' Add a table.
    Set objDBTables = objReport.Database.Tables
    objReport.Database.Tables.Add "Northwind.dbo.Customers", "", , , "pdsodbc.dll", _
        "SQLServerNorthwind.dsn", "", "dbo", , ""
    objReport.Database.Tables(1).Name = "Customers"
    
    ' Add a table.
    Set objDBTables = objReport.Database.Tables
    objReport.Database.Tables.Add "Northwind.dbo.Orders", "", , , "pdsodbc.dll", _
        "SQLServerNorthwind.dsn", "", "dbo", , ""
    objReport.Database.Tables(2).Name = "Orders"
    
    
    ' Set database link.
    Set objOneDBTable = objReport.Database.Tables.Item(1)
    Set objManyDBTable = objReport.Database.Tables.Item(2)
    Set objOneDBField = objReport.Database.Tables.Item(1).Fields.Item(1)
    Set objManyDBField = objReport.Database.Tables.Item(2).Fields.Item(2)
    Set objDBlinks = objReport.Database.Links
    Set objDBLink = objDBlinks.Add(objOneDBTable, objManyDBTable, objOneDBField, _
        objManyDBField, crJTEqual, crLTLookupParallel, False, True)
    
    
    ' Set the report options.
    With objReport
        .CaseInsensitiveSQLData = True
        .ConvertDateTimeType = crKeepDateTimeType
        .ConvertNullFieldToDefault = False
        .MorePrintEngineErrorMessages = True
        .TranslateDosMemos = True
        .TranslateDosStrings = True
        .CaseInsensitiveSQLData = True
        .PerformGroupingOnServer = False
        .UseIndexForSpeed = True
        .EnableAsyncQuery = False
        .EnableGeneratingDataForHiddenObject = False
        .EnableParameterPrompting = False
        .EnablePerformQueriesAsynchronously = False
        .EnableSelectDistinctRecords = False
        .VerifyOnEveryPrint = False
        .PrintDate = DateSerial(2002, 5, 3)
        .SavePreviewPicture = False
        .BottomMargin = 262
        .LeftMargin = 259
        .RightMargin = 262
        .TopMargin = 259
    End With
    
        
    ' Add group Object. - "{Customers.CompanyName}"
    objReport.AddGroup 0, objReport.Database.Tables.Item(1).Fields(3), _
        crGCAnyValue, crAscendingOrder
    
    ' Set group options.
    With objReport.Areas("GH1")
        .Suppress = False
        .HideForDrillDown = False
        .DiscardOtherGroups = False
        .KeepGroupTogether = False
        .RepeatGroupHeader = False
        .EnableHierarchicalGroupSorting = False
        .GroupIndent = 0
    End With
    
    ' Set area format for Report Header.
    With objReport.Areas("RH")
        .Suppress = False
        .NewPageAfter = False
        .NewPageBefore = True
        .KeepTogether = False
        .HideForDrillDown = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .PrintAtBottomOfPage = False
    End With
    
    ' Set section format for Report Header a.
    With objReport.Sections("RHa")
        .Height = 220
        .Suppress = False
        .BackColor = vbWhite
        .KeepTogether = True
        .NewPageAfter = False
        .NewPageBefore = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .SuppressIfBlank = False
        .UnderlaySection = False
        .ResetPageNumberAfter = False
    End With
    
    ' Set area format for Page Header.
    With objReport.Areas("PH")
        .Suppress = False
        .NewPageAfter = False
        .NewPageBefore = False
        .KeepTogether = True
        .HideForDrillDown = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .PrintAtBottomOfPage = False
    End With
    
    ' Set section format for Page Header a.
    With objReport.Sections("PHa")
        .Height = 927
        .Suppress = False
        .BackColor = vbWhite
        .KeepTogether = True
        .NewPageAfter = False
        .NewPageBefore = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .SuppressIfBlank = False
        .UnderlaySection = False
        .ResetPageNumberAfter = False
    End With
    
    
    ' Set area format for Group Header #1.
    With objReport.Areas("GH1")
        .Suppress = False
        .NewPageAfter = False
        .NewPageBefore = False
        .KeepTogether = False
        .HideForDrillDown = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .PrintAtBottomOfPage = False
    End With
    
    ' Set section format for Group Header #1a.
    With objReport.Sections("GH1a")
        .Height = 230
        .Suppress = False
        .BackColor = vbWhite
        .KeepTogether = True
        .NewPageAfter = False
        .NewPageBefore = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .SuppressIfBlank = False
        .UnderlaySection = False
        .ResetPageNumberAfter = False
    End With
        
    
    ' Set area format for Details.
    With objReport.Areas("D")
        .Suppress = False
        .NewPageAfter = False
        .NewPageBefore = False
        .KeepTogether = False
        .HideForDrillDown = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .PrintAtBottomOfPage = False
    End With
    
    ' Set section format for Details a.
    With objReport.Sections("Da")
        .Height = 230
        .Suppress = False
        .BackColor = vbWhite
        .KeepTogether = True
        .NewPageAfter = False
        .NewPageBefore = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .SuppressIfBlank = False
        .UnderlaySection = False
        .ResetPageNumberAfter = False
    End With
    
    ' Set area format for Group Footer #1.
    With objReport.Areas("GF1")
        .Suppress = False
        .NewPageAfter = False
        .NewPageBefore = False
        .KeepTogether = False
        .HideForDrillDown = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .PrintAtBottomOfPage = False
    End With
    
    ' Set section format for Group Footer #1a.
    With objReport.Sections("GF1a")
        .Height = 240
        .Suppress = False
        .BackColor = vbWhite
        .KeepTogether = True
        .NewPageAfter = False
        .NewPageBefore = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .SuppressIfBlank = False
        .UnderlaySection = False
        .ResetPageNumberAfter = False
    End With
    
    ' Set area format for Group Footer #1.
    With objReport.Areas("GF1")
        .Suppress = False
        .NewPageAfter = False
        .NewPageBefore = False
        .KeepTogether = False
        .HideForDrillDown = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .PrintAtBottomOfPage = False
    End With
    
    ' Set area format for Page Footer.
    With objReport.Areas("PF")
        .Suppress = False
        .NewPageAfter = False
        .NewPageBefore = False
        .KeepTogether = True
        .HideForDrillDown = False
        .PrintAtBottomOfPage = True
        .ResetPageNumberAfter = False
        .PrintAtBottomOfPage = True
    End With
    
    ' Set section height for Page Footer a.
    objReport.Sections("PFa").Height = 690
    
    ' Set area format for Report Footer.
    With objReport.Areas("RF")
        .Suppress = False
        .NewPageAfter = False
        .NewPageBefore = False
        .KeepTogether = False
        .HideForDrillDown = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .PrintAtBottomOfPage = False
    End With
    
    ' Set section format for Report Footer a.
    With objReport.Sections("RFa")
        .Height = 294
        .Suppress = False
        .BackColor = vbWhite
        .KeepTogether = True
        .NewPageAfter = False
        .NewPageBefore = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .SuppressIfBlank = False
        .UnderlaySection = False
        .ResetPageNumberAfter = False
    End With
    
    ' Add a text object.
    Set objTextObject = objReport.Sections("PHa").AddTextObject("Company", 60, 697)
    
    With objTextObject
        .Height = 230
        .Width = 2648
        
        .TextColor = vbBlack
        .Font.Italic = False
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Strikethrough = False
        .Font.Underline = True
        .Font.Bold = True
        
        .Suppress = False
        .HorAlignment = crLeftAlign
        .KeepTogether = True
        .CanGrow = False
        .LeftLineStyle = crLSNoLine
        .RightLineStyle = crLSNoLine
        .TopLineStyle = crLSNoLine
        .BottomLineStyle = crLSNoLine
        .BackColor = vbWhite
        .SuppressIfDuplicated = False
    End With
    
    Set objFont = objTextObject.Font
    
    
    ' Add a text object.
    Set objTextObject = objReport.Sections("PHa").AddTextObject("Contact", 2828, 697)
    
    With objTextObject
        .Height = 230
        .Width = 2036
        .TextColor = vbBlack
        Set .Font = objFont
        .Suppress = False
        .HorAlignment = crLeftAlign
        .KeepTogether = True
        .CanGrow = False
        .LeftLineStyle = crLSNoLine
        .RightLineStyle = crLSNoLine
        .TopLineStyle = crLSNoLine
        .BottomLineStyle = crLSNoLine
        .BackColor = vbWhite
        .SuppressIfDuplicated = False
    End With
    

    ' Add a text object.
    Set objTextObject = objReport.Sections("PHa").AddTextObject("Required Date", 5230, 697)
    
    With objTextObject
        .Height = 230
        .Width = 1390
        .TextColor = vbBlack
        Set .Font = objFont
        .Suppress = False
        .HorAlignment = crLeftAlign
        .KeepTogether = True
        .CanGrow = False
        .LeftLineStyle = crLSNoLine
        .RightLineStyle = crLSNoLine
        .TopLineStyle = crLSNoLine
        .BottomLineStyle = crLSNoLine
        .BackColor = vbWhite
        .SuppressIfDuplicated = False
    End With
    
    ' Add a text object.
    Set objTextObject = objReport.Sections("PHa").AddTextObject("Freight", 6867, 690)
    
    With objTextObject
        .Height = 230
        .Width = 977
        .TextColor = vbBlack
        Set .Font = objFont
        .Suppress = False
        .HorAlignment = crRightAlign
        .KeepTogether = True
        .CanGrow = False
        .LeftLineStyle = crLSNoLine
        .RightLineStyle = crLSNoLine
        .TopLineStyle = crLSNoLine
        .BottomLineStyle = crLSNoLine
        .BackColor = vbWhite
        .SuppressIfDuplicated = False
    End With
    
    ' Add a special variable field object to Page Header a.

    Set objSpecialField = objReport.Sections("PHa").AddSpecialVarFieldObject(crSVTPrintDate, 60, 230)
    
    With objSpecialField
        .Height = 230
        .Width = 799
    
        .TextColor = vbBlack
        .Font.Italic = False
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Strikethrough = False
        .Font.Underline = False
        .Font.Bold = False
    
        .Suppress = False
        .HorAlignment = crDefaultAlign
        .KeepTogether = True
        .CanGrow = False
        .LeftLineStyle = crLSNoLine
        .RightLineStyle = crLSNoLine
        .TopLineStyle = crLSNoLine
        .BottomLineStyle = crLSNoLine
        .BackColor = vbWhite
        .SuppressIfDuplicated = False
        .UseSystemDefaults = True
        .DateWindowsDefaultType = crUseWindowsShortDate
        .DateOrder = crMonthDayYear
        .YearType = crLongYear
        .MonthType = crNumericMonth
        .DayType = crNumericDay
        .LeadingDayType = crNoLeadingDay
        .DateFirstSeparator = "/"
        .DateSecondSeparator = "/"
        .LeadingDaySeparator = ""
    End With
    
    Set objFont = objSpecialField.Font
    
    
    ' Add a text object.
    Set objTextObject = objReport.Sections("PHa").AddTextObject("Freight costs across Orders", 3840, 187)
    
    With objTextObject
        .Height = 300
        .Width = 3720
        .TextColor = vbBlack
        Set .Font = objFont
        .Suppress = False
        .HorAlignment = crHorCenterAlign
        .KeepTogether = True
        .CanGrow = False
        .LeftLineStyle = crLSNoLine
        .RightLineStyle = crLSNoLine
        .TopLineStyle = crLSNoLine
        .BottomLineStyle = crLSNoLine
        .BackColor = vbWhite
        .SuppressIfDuplicated = False
    End With
    
    ' Add a database field object to Group Header #1a.
    Set objFieldObject = objReport.Sections("GH1a").AddFieldObject("{Customers.CompanyName}", 60, 0)
    
    With objFieldObject
        .Height = 230
        .Width = 2648
        .TextColor = vbBlack
        Set .Font = objFont
        .Suppress = False
        .HorAlignment = crDefaultAlign
        .KeepTogether = True
        .CanGrow = False
        .LeftLineStyle = crLSNoLine
        .RightLineStyle = crLSNoLine
        .TopLineStyle = crLSNoLine
        .BottomLineStyle = crLSNoLine
        .BackColor = vbWhite
        .SuppressIfDuplicated = False
        .UseSystemDefaults = True
    End With
    
    ' Add a database field object to Group Header #1a.
    Set objFieldObject = objReport.Sections("GH1a").AddFieldObject("{Customers.ContactName}", 2828, 0)
    
    With objFieldObject
        .Height = 230
        .Width = 2036
        .TextColor = vbBlack
        Set .Font = objFont
        .Suppress = False
        .HorAlignment = crDefaultAlign
        .KeepTogether = True
        .CanGrow = False
        .LeftLineStyle = crLSNoLine
        .RightLineStyle = crLSNoLine
        .TopLineStyle = crLSNoLine
        .BottomLineStyle = crLSNoLine
        .BackColor = vbWhite
        .SuppressIfDuplicated = False
        .UseSystemDefaults = True
    End With
    
    
    ' Add a database field object to Details a.
    Set objFieldObject = objReport.Sections("Da").AddFieldObject("{Orders.RequiredDate}", 5290, 0)
    
    With objFieldObject
        .TextColor = vbBlack
        Set .Font = objFont
        .Height = 230
        .Width = 1000
        .Suppress = False
        .HorAlignment = crDefaultAlign
        .KeepTogether = True
        .CanGrow = False
        .LeftLineStyle = crLSNoLine
        .RightLineStyle = crLSNoLine
        .TopLineStyle = crLSNoLine
        .BottomLineStyle = crLSNoLine
        .BackColor = vbWhite
        .SuppressIfDuplicated = False
        .UseSystemDefaults = False
        .DateWindowsDefaultType = crNotUsingWindowsDefaults
        .DateOrder = crMonthDayYear
        .YearType = crLongYear
        .MonthType = crLeadingZeroNumericMonth
        .DayType = crLeadingZeroNumericDay
        .LeadingDayType = crNoLeadingDay
        .DateFirstSeparator = "/"
        .DateSecondSeparator = "/"
        .LeadingDaySeparator = ""
        .HourType = crNoHour
        .MinuteType = crNoMinute
        .SecondType = crNumericNoSecond
    End With
    
    ' Add a database field object to Details a.
    Set objFieldObject = objReport.Sections("Da").AddFieldObject("{Orders.Freight}", 6960, 0)
    
    With objFieldObject
        .TextColor = vbBlack
        Set .Font = objFont
        .Height = 230
        .Width = 981
        .Suppress = False
        .HorAlignment = crDefaultAlign
        .KeepTogether = True
        .CanGrow = False
        .LeftLineStyle = crLSNoLine
        .RightLineStyle = crLSNoLine
        .TopLineStyle = crLSNoLine
        .BottomLineStyle = crLSNoLine
        .BackColor = vbWhite
        .SuppressIfDuplicated = False
        .UseSystemDefaults = True
        .SuppressIfZero = False
        .NegativeType = crBracketed
        .ThousandsSeparators = True
        .UseLeadingZero = True
        .DecimalPlaces = 2
        .RoundingType = crRoundToHundredth
        .CurrencySymbolType = crCSTFloatingSymbol
        .UseOneSymbolPerPage = False
        .CurrencyPositionType = crLeadingCurrencyOutsideNegative
        .ThousandSymbol = ","
        .DecimalSymbol = "."
        .CurrencySymbol = "$"
    End With
    
    ' Add a summary field object to Group Footer #1a.
    Set objFieldObject = objReport.Sections("GF1a").AddSummaryFieldObject("{Orders.Freight}", crSTSum, 6960, 0)
    
    With objFieldObject
        .TextColor = vbBlack
        Set .Font = objFont
        .Height = 240
        .Width = 986
        .Suppress = False
        .HorAlignment = crDefaultAlign
        .KeepTogether = True
        .CanGrow = False
        .LeftLineStyle = crLSNoLine
        .RightLineStyle = crLSNoLine
        .TopLineStyle = crLSNoLine
        .BottomLineStyle = crLSNoLine
        .BackColor = vbWhite
        .SuppressIfDuplicated = False
        .UseSystemDefaults = True
        .SuppressIfZero = False
        .NegativeType = crBracketed
        .ThousandsSeparators = True
        .UseLeadingZero = True
        .DecimalPlaces = 2
        .RoundingType = crRoundToHundredth
        .CurrencySymbolType = crCSTFloatingSymbol
        .UseOneSymbolPerPage = False
        .CurrencyPositionType = crLeadingCurrencyOutsideNegative
        .ThousandSymbol = ","
        .DecimalSymbol = "."
        .CurrencySymbol = "$"
    End With
    
    
    ' Add a special variable field object to Page Footer a.
    Set objSpecialField = objReport.Sections("PFa").AddSpecialVarFieldObject(crSVTPageNumber, 10405, 460)
    
    With objSpecialField
        .Height = 230
        .Width = 974
        .TextColor = vbBlack
        Set .Font = objFont
        .Suppress = False
        .HorAlignment = crDefaultAlign
        .KeepTogether = True
        .CanGrow = False
        .LeftLineStyle = crLSNoLine
        .RightLineStyle = crLSNoLine
        .TopLineStyle = crLSNoLine
        .BottomLineStyle = crLSNoLine
        .BackColor = vbWhite
        .SuppressIfDuplicated = False
        .UseSystemDefaults = True
        .SuppressIfZero = False
        .NegativeType = crLeadingMinus
        .ThousandsSeparators = True
        .UseLeadingZero = True
        .DecimalPlaces = 0
        .RoundingType = crRoundToUnit
        .CurrencySymbolType = crCSTNoSymbol
        .UseOneSymbolPerPage = False
        .CurrencyPositionType = crLeadingCurrencyInsideNegative
        .ThousandSymbol = ","
        .DecimalSymbol = "."
        .CurrencySymbol = ""
    End With
    
    ' Add a summary field object to Report Footer a.
    Set objFieldObject = objReport.Sections("RFa").AddSummaryFieldObject("{Orders.Freight}", crSTSum, 6697, 0)
    
    With objFieldObject
        .Height = 294
        .Width = 1305
        .TextColor = vbBlack
        Set .Font = objFont
        .Suppress = False
        .HorAlignment = crDefaultAlign
        .KeepTogether = True
        .CanGrow = False
        .LeftLineStyle = crLSNoLine
        .RightLineStyle = crLSNoLine
        .TopLineStyle = crLSNoLine
        .BottomLineStyle = crLSNoLine
        .BackColor = vbWhite
        .SuppressIfDuplicated = False
        .UseSystemDefaults = True
        .SuppressIfZero = False
        .NegativeType = crBracketed
        .ThousandsSeparators = True
        .UseLeadingZero = True
        .DecimalPlaces = 2
        .RoundingType = crRoundToHundredth
        .CurrencySymbolType = crCSTFloatingSymbol
        .UseOneSymbolPerPage = False
        .CurrencyPositionType = crLeadingCurrencyOutsideNegative
        .ThousandSymbol = ","
        .DecimalSymbol = "."
        .CurrencySymbol = "$"
    End With
    
    objReport.Save ReportName
    
    ' Close the report.
    Set CreateReport = objReport
    
End Function


