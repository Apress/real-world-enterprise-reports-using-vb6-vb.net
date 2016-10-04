Private Sub objListBox_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim objListbox As ListBox
    Dim objButton As Button
    Dim iIndex As Integer

    For Each objButton In objListBoxButtonColl

        iIndex += 1

        If objButton Is sender Then
            Exit For
        End If

    Next

    objListbox = CType(objListBoxColl(iIndex), ListBox)

    objListbox.ClearSelection()

End Sub
