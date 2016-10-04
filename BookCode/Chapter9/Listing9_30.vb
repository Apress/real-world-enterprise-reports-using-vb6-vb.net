Private Sub cmdOK_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cmdOK.Click

    Dim objSQL As New StringBuilder()
    Dim cSQL As String
    Dim objCriteria As New StringBuilder()
    Dim cCriteria As String
    Dim cDateFrom As String
    Dim cDateTo As String

    cDateFrom = objDateFromColl(SELECT_ENTER_DATE).Text
    cDateTo = objDateToColl(SELECT_ENTER_DATE).Text

    With objSQL
        .AppendFormat (GetDateSQL("createdate", cDateFrom, cDateTo))
        .AppendFormat (GetListBoxSQL(objListBoxColl(SELECT_PRODUCT), "ProductID"))
        .AppendFormat (GetListBoxSQL(objListBoxColl(SELECT_SOURCE), "SourceID"))
        .AppendFormat (GetComboBoxSQL(SELECT_DEPARTMENT, "deptid"))
        .AppendFormat (GetCheckBoxSQL(SELECT_COMPLEX_QUESTION, "questiontype"))
    End With

    With objCriteria
        .AppendFormat (GetDateCriteria("Create Date", cDateFrom, cDateTo))
        .AppendFormat (GetListBoxCriteria(objListBoxColl(SELECT_PRODUCT), "Product"))
        .AppendFormat (GetListBoxCriteria(objListBoxColl(SELECT_SOURCE), "Source"))
        .AppendFormat (GetComboBoxCriteria(SELECT_DEPARTMENT, "Department"))
        .AppendFormat (GetCheckBoxCriteria(SELECT_COMPLEX_QUESTION, "Complex Question"))
    End With

    cSQL = objSQL.ToString
    cCriteria = objCriteria.ToString

    If objSQL.ToString <> String.Empty Then
        cSQL = cSQL.Substring(0, cSQL.Length - 5)
    End If

    If objCriteria.ToString <> String.Empty Then
        cCriteria = cCriteria.Substring(0, cCriteria.Length - 5)
    End If

End Sub
