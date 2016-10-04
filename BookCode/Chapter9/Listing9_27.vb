Private Sub ShowListBox(ByVal iIndex As Short, ByVal iLeft As Short, _
    ByVal iTop As Short, ByVal iWidth As Short, ByVal iHeight As Short, _
    ByVal cCaption As String)

    Dim iListBoxHeight As Short

    Call AddDynamicListBoxLabel(iIndex, iLeft, iTop, iWidth, cCaption)

    Call AddDynamicListBox(iIndex, iLeft, iTop + iHeight + 5, iWidth, 180)

    iListBoxHeight = Replace(objListBoxColl(iIndex).style("height"), "px", String.Empty)

    Call AddDynamicListBoxButton(iIndex, iLeft, (iTop + iHeight + 5) + _
         iListBoxHeight + 5, iWidth, cCaption)

End Sub
