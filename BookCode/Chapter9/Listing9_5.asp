Function GetListBoxSQL(cValues, cColumn) 
    Dim cSQL     
    Dim aValues
    Dim x
    Dim cIDs
    
    aValues = Split(cValues, ",")
    
    For x = 0 to Ubound(aValues) Step 2
          cIDs = cIDs & aValues(x) & ", "
    Next

    cIDs = "(" & Mid(cIDs, 1, Len(cIDs) - 2) & ")"
    
    If cIDs <> "" Then
          cSQL = cColumn & " IN " & cIDs & " AND "
    End if    
       
    GetListBoxSQL = cSQL
    
End Function

Function GetListBoxCriteria(cValues, cDescr) 
    Dim cResult 
    Dim cNames 
    Dim aValues
    Dim x
    
    aValues = Split(cValues, ",")
    
    For x = 1 to Ubound(aValues) Step 2
          cNames = cNames & aValues(x) & ", "
    Next
    
    If cNames <> "" Then
      cNames = "(" & Mid(cNames, 1, Len(cNames) - 2) & ")"

      cResult = cResult & "the " & cDescr & _
          " is among " & cNames & " and " 
    End if   

    GetListBoxCriteria = cResult
    
End Function
