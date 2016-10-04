Public Enum ExportFormats
    Excel = 0
    PDF = 1
    HTML = 2
    RTF = 3
    ASCII = 4
    XML = 5
End Enum

<WebMethod()> Public Sub EmployeeRpt(ByVal cWhere As String, _
    ByVal iExport As ExportFormats)

    Dim objSQL As New System.Text.StringBuilder()

    objSQL.Append ("SELECT * " & _
                  "FROM Employee " & _
                  "WHERE ")

    objSQL.Append (cWhere)

    'Execute report code using your reporting tool and
    'dump the output to the export format of choice

End Sub
