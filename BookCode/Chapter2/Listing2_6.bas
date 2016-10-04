Function RunCommand(oConn As ADODB.Connection, cSQL As String, _
        iCommandType As Integer, iCursorType As Integer, _
        Optional vntColumn As Variant, _
        Optional vntID As Variant) As ADODB.Recordset
    
    Dim oRS As New ADODB.Recordset
    Dim oCmd As New ADODB.Command

    oConn.CursorLocation = adUseClient
    
    With oCmd
    
        Set .ActiveConnection = oConn
        
        .CommandText = cSQL
        .CommandType = iCommandType
        .CommandTimeout = 1200
        
        If iCommandType = adCmdStoredProc And _
            Not IsMissing(vntColumn) And _
            Not IsMissing(vntID) Then
                    
            If IsNumeric(vntID) Then
                oCmd.Parameters.Append _
                oCmd.CreateParameter(vntColumn, adInteger, adParamInput, 9, vntID)
            Else
                oCmd.Parameters.Append _
                oCmd.CreateParameter(vntColumn, adVarChar, adParamInput, 50, vntID)
            End If
            
        End If
    
    End With
    
    With oRS
        .CursorType = iCursorType
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .Open oCmd
    End With

    Set RunCommand = oRS
    
End Function
