Private Sub AddDynamicListBox(ByVal iIndex As Short, ByVal iLeft As Short, _
    ByVal iTop As Short, ByVal iWidth As Short, ByVal iHeight As Short)

    Dim objListBox As New ListBox()

    With objListBox
        .Style("position") = "absolute"
        .Style("left") = iLeft & "px"
        .Style("top") = iTop & "px"
        .Style("height") = iHeight & "px"
        .Style("width") = iWidth & "px"
        .SelectionMode = ListSelectionMode.Multiple
    End With

    objListBoxColl.Add (objListBox)

    Panel1.Controls.Add (objListBox)

End Sub
