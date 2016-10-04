Function GetListBoxSQL(ByRef objListBox As ListBox, _
    ByVal cColumn As String) As String

    Dim cSQL As String = String.Empty
    Dim cList As String

    cList = ParseIt(objListBox, True, False, 0)

    If cList <> String.Empty Then
        cSQL = cColumn & " IN " & cList & " AND "
    End If

    GetListBoxSQL = cSQL

End Function
