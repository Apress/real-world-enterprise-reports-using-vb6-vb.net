Function ParseIt(ByVal oList As ListBox, ByVal bTagged As Boolean, _
    ByVal bQuotes As Boolean, ByVal iCol As Short) As String

    Dim objResult As New StringBuilder("(")
    Dim cResult As String = String.Empty
    Dim cQuotes As String
    Dim cData As String
    Dim oTemp As ListItem
    Dim oCollection As Object

    If bQuotes Then
        cQuotes = ControlChars.Quote
    Else
        cQuotes = String.Empty
    End If

    If bTagged Then

        For Each oTemp In oList.Items

            If oTemp.Selected Then

                If iCol = 0 Then
                    cData = oTemp.Value
                Else
                    cData = oTemp.Text
                End If

                objResult.AppendFormat (cQuotes & cData & cQuotes & ",")

            End If

        Next

    Else

        For Each oTemp In oList.Items

            If iCol = 0 Then
                cData = oTemp.Value
            Else
                cData = oTemp.Text
            End If

            objResult.AppendFormat (cQuotes & cData & cQuotes & ",")

        Next

    End If

    cResult = objResult.ToString

    If bQuotes Then
        cResult = Mid$(cResult, 1, Len(cResult) - 2) & cQuotes & ")"
    Else
        cResult = Mid$(cResult, 1, Len(cResult)) & cQuotes & ")"
    End If

    cResult = Replace(cResult, ",)", ")")

    cResult = Replace(cResult, ",,", String.Empty)

    If cResult = "()" Then
        cResult = String.Empty
    End If

    ParseIt = cResult

End Function
