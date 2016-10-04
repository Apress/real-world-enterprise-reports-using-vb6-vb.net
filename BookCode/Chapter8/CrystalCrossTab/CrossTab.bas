Attribute VB_Name = "Module1"
Option Explicit

Sub main()
    Call CreateReport("C:\Reports\CrossTab.rpt")
End Sub

Function CreateReport(cReportName As String) As Boolean
    Dim objApplication As CRAXDRT.Application
    Dim objReport As CRAXDRT.Report
    Dim objDBTables As CRAXDRT.DatabaseTables
    Dim objSourceDBTable As CRAXDRT.DatabaseTable
    Dim objDestDBTable As CRAXDRT.DatabaseTable
    Dim objSourceDBField As CRAXDRT.DatabaseFieldDefinition
    Dim objDestDBField As CRAXDRT.DatabaseFieldDefinition
    Dim objDBlinks As CRAXDRT.TableLinks
    Dim objDBLink As CRAXDRT.TableLink
    Dim objCrossTabObject As CRAXDRT.CrossTabObject
    Dim objCrossTabGroup As CRAXDRT.CrossTabGroup
    Dim objFieldObject As CRAXDRT.FieldObject
    Dim cDatabasePath As String
    
    cDatabasePath = "C:\Program Files\Seagate Software\Crystal Reports" & _
                    "\Samples\En\Databases\xtreme.mdb"
    
    Set objApplication = New CRAXDRT.Application
    Set objReport = objApplication.NewReport
    
    With objReport
        Set objDBTables = .Database.Tables
        
        .Database.Tables.Add cDatabasePath, "Customer"
        .Database.Tables(1).SetSessionInfo "Admin", ""
    
        .Database.Tables.Add cDatabasePath, "Orders"
        .Database.Tables(2).SetSessionInfo "Admin", ""
    
        .Database.Tables.Add cDatabasePath, "Orders Detail"
        .Database.Tables(3).SetSessionInfo "Admin", ""
    
        .Database.Tables.Add cDatabasePath, "Product"
        .Database.Tables(4).SetSessionInfo "Admin", ""
            
        Set objSourceDBTable = .Database.Tables.Item(3) 'Orders Detail
        Set objDestDBTable = .Database.Tables.Item(4) 'Product
        Set objSourceDBField = .Database.Tables.Item(3).Fields.Item(2) '{Orders Detail.Product ID}
        Set objDestDBField = .Database.Tables.Item(4).Fields.Item(1) '{Product.Product ID}
        Set objDBlinks = .Database.Links
        Set objDBLink = objDBlinks.Add(objSourceDBTable, objDestDBTable, objSourceDBField, _
            objDestDBField, crJTEqual, crLTLookupParallel, False, True)
    
        Set objSourceDBTable = .Database.Tables.Item(3) 'Orders Detail
        Set objDestDBTable = .Database.Tables.Item(2) 'Orders
        Set objSourceDBField = .Database.Tables.Item(3).Fields.Item(1) '{Orders Detail.Order ID}
        Set objDestDBField = .Database.Tables.Item(2).Fields.Item(1) '{Orders.Order ID}
        Set objDBlinks = .Database.Links
        Set objDBLink = objDBlinks.Add(objSourceDBTable, objDestDBTable, objSourceDBField, _
            objDestDBField, crJTEqual, crLTLookupParallel, False, True)
    
        Set objSourceDBTable = .Database.Tables.Item(2) 'Orders
        Set objDestDBTable = .Database.Tables.Item(1) 'Customer
        Set objSourceDBField = .Database.Tables.Item(2).Fields.Item(3) '{Orders.Customer ID}
        Set objDestDBField = .Database.Tables.Item(1).Fields.Item(1) '{Customer.Customer ID}
        Set objDBlinks = .Database.Links
        Set objDBLink = objDBlinks.Add(objSourceDBTable, objDestDBTable, objSourceDBField, _
            objDestDBField, crJTEqual, crLTLookupParallel, False, True)
    
        ' Set the report options
        .PaperOrientation = crLandscape
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
        .SavePreviewPicture = False
        .BottomMargin = 262
        .LeftMargin = 259
        .RightMargin = 262
        .TopMargin = 259
        
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

    ' Set section height for Report Header a.
    With objReport.Sections("RHa")
        .Height = 1680
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

    ' Set section height for Page Header a.
    With objReport.Sections("PHa")
        .Height = 460
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
        .Suppress = True
        .NewPageAfter = False
        .NewPageBefore = False
        .KeepTogether = False
        .HideForDrillDown = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .PrintAtBottomOfPage = False
    End With


    ' Set section height for Details a.
    With objReport.Sections("Da")
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
        .Suppress = True
        .NewPageAfter = False
        .NewPageBefore = False
        .KeepTogether = False
        .HideForDrillDown = False
        .PrintAtBottomOfPage = False
        .ResetPageNumberAfter = False
        .PrintAtBottomOfPage = False
    End With

    ' Set section height for Report Footer a.
    With objReport.Sections("RFa")
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
    
    
    Set objCrossTabObject = objReport.Sections("RHa").AddCrossTabObject(0, 0)
    
    With objCrossTabObject
        .SummaryFields.Add "{Orders.Order Amount}"
        .EnableRepeatRowLabels = True
        .EnableSuppressColumnGrandTotals = False
        .EnableSuppressRowGrandTotals = False
        .RowGrandTotalColor = vbRed
        .ColumnGrandTotalColor = vbGreen
        .Suppress = False
        .KeepTogether = False
        .LeftLineStyle = crLSNoLine
        .RightLineStyle = crLSNoLine
        .TopLineStyle = crLSNoLine
        .BottomLineStyle = crLSNoLine
        .BackColor = vbWhite
    End With
    
    Set objCrossTabGroup = objCrossTabObject.RowGroups.Add("{Customer.Region}")

    With objCrossTabGroup
        .BackColor = vbWhite
        .Condition = crGCAnyValue
        .SortDirection = crAscendingOrder
    End With

    Set objCrossTabGroup = objCrossTabObject.ColumnGroups.Add("{Product.Product Name}")

    With objCrossTabGroup
        .BackColor = vbWhite
        .Condition = crGCAnyValue
        .SortDirection = crAscendingOrder
    End With
    
    
    
'    ' Add a special variable field object to Page Header a.
'    Set objFieldObject = objReport.Sections("PHa").AddSpecialVarFieldObject(crSVTPrintDate, 0, 230)
'
'    With objFieldObject
'        .Height = 230
'        .Width = 720
'        .TextColor = vbBlack
'        .Font.Italic = False
'        .Font.Name = "Times New Roman"
'        .Font.Size = 10
'        .Font.Strikethrough = False
'        .Font.Underline = False
'        .Font.Bold = False
'        .Suppress = False
'        .HorAlignment = crDefaultAlign
'        .KeepTogether = True
'        .CanGrow = False
'        .LeftLineStyle = crLSNoLine
'        .RightLineStyle = crLSNoLine
'        .TopLineStyle = crLSNoLine
'        .BottomLineStyle = crLSNoLine
'        .BackColor = vbWhite
'        .SuppressIfDuplicated = False
'        .UseSystemDefaults = True
'        .DateWindowsDefaultType = crUseWindowsShortDate
'        .DateOrder = crMonthDayYear
'        .YearType = crLongYear
'        .MonthType = crNumericMonth
'        .DayType = crNumericDay
'        .LeadingDayType = crNoLeadingDay
'        .DateFirstSeparator = "/"
'        .DateSecondSeparator = "/"
'        .LeadingDaySeparator = ""
'    End With
'
'    ' Add a special variable field object to Page Footer a.
'    Set objFieldObject = objReport.Sections("PFa").AddSpecialVarFieldObject(crSVTPageNumber, 0, 460)
'
'    With objFieldObject
'        .Height = 230
'        .Width = 878
'        .TextColor = vbBlack
'        .Font.Italic = False
'        .Font.Name = "Times New Roman"
'        .Font.Size = 10
'        .Font.Strikethrough = False
'        .Font.Underline = False
'        .Font.Bold = False
'        .Suppress = False
'        .HorAlignment = crDefaultAlign
'        .KeepTogether = True
'        .CanGrow = False
'        .LeftLineStyle = crLSNoLine
'        .RightLineStyle = crLSNoLine
'        .TopLineStyle = crLSNoLine
'        .BottomLineStyle = crLSNoLine
'        .BackColor = vbWhite
'        .SuppressIfDuplicated = False
'        .UseSystemDefaults = True
'        .SuppressIfZero = False
'        .NegativeType = crLeadingMinus
'        .ThousandsSeparators = True
'        .UseLeadingZero = True
'        .DecimalPlaces = 0
'        .RoundingType = crRoundToUnit
'        .CurrencySymbolType = crCSTNoSymbol
'        .UseOneSymbolPerPage = False
'        .CurrencyPositionType = crLeadingCurrencyInsideNegative
'        .ThousandSymbol = ","
'        .DecimalSymbol = "."
'        .CurrencySymbol = ""
'    End With
    
    objReport.Save cReportName

End Function

