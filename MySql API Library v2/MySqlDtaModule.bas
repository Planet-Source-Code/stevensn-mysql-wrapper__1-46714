Attribute VB_Name = "MySqlDtaModule"
Option Explicit

Sub DefErr(Number As Long, Source, Description)
    Err.Raise Number, Source, Description
End Sub

Function RemoveLastChar(lpStr As String, lpChar As String) As String
    Dim strTemp As String
    Dim intLoopCounter As Long
    Dim intNumCycle As Long
    
    intNumCycle = 0
    
    strTemp = lpStr
    
    For intLoopCounter = Len(strTemp) To 1 Step -1
        intNumCycle = intNumCycle + 1
        If intNumCycle > 4 Then
            RemoveLastChar = strTemp
            Exit For
        End If
        
        If Mid(strTemp, intLoopCounter, 1) = lpChar Then
            If intLoopCounter = 1 Then
                RemoveLastChar = ""
            Else
                RemoveLastChar = left(strTemp, intLoopCounter - 1)
            End If
        End If
    Next
End Function

Function StruFieldToQuery(pField As MySqlField) As String
    Dim strSql As String
    
    If pField.FieldType = fieldtypeBlob Then
        If Not pField.FieldNull Then
            strSql = "BLOB NOT NULL"
        Else
            strSql = "BLOB"
        End If
    End If
    
    If pField.FieldType = fieldtypeDate Or pField.FieldType = fieldtypeNewDate Then
        If Not pField.FieldNull Then
            If pField.FieldDefault = "" Then
                strSql = "DATE NOT NULL DEFAULT '0000-00-00'"
            Else
                strSql = "DATE NOT NULL DEFAULT '" & pField.FieldDefault & "'"
            End If
        Else
            If pField.FieldDefault = "" Then
                strSql = "DATE"
            Else
                strSql = "DATE DEFAULT '" & pField.FieldDefault & "'"
            End If
        End If
    End If
    
    If pField.FieldType = fieldtypeDateTime Then
        If Not pField.FieldNull Then
            If pField.FieldDefault = "" Then
                strSql = "DATETIME NOT NULL DEFAULT '0000-00-00 00:00:00'"
            Else
                strSql = "DATETIME NOT NULL DEFAULT '" & pField.FieldDefault & "'"
            End If
        Else
            If pField.FieldDefault = "" Then
                strSql = "DATETIME"
            Else
                strSql = "DATETIME DEFAULT '" & pField.FieldDefault & "'"
            End If
        End If
    End If
    
    If pField.FieldType = fieldtypeDouble Or pField.FieldType = fieldtypeDecimal Then
        If Not pField.FieldNull Then
            If pField.FieldDefault = "" Then
                strSql = "DECIMAL(" & pField.FieldSize & "," & pField.FieldDecimals & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & " NOT NULL " & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '0'"
            Else
                strSql = "DECIMAL(" & pField.FieldSize & "," & pField.FieldDecimals & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & " NOT NULL " & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '" & pField.FieldDefault & "'"
            End If
        Else
            If pField.FieldDefault = "" Then
                strSql = "DECIMAL(" & pField.FieldSize & "," & pField.FieldDecimals & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "")
            Else
                strSql = "DECIMAL(" & pField.FieldSize & "," & pField.FieldDecimals & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '" & pField.FieldDefault & "'"
            End If
        End If
    End If

    If pField.FieldType = fieldtypeFloat Then
        If Not pField.FieldNull Then
            If pField.FieldDefault = "" Then
                strSql = "FLOAT(" & pField.FieldSize & "," & pField.FieldDecimals & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & " NOT NULL " & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '0'"
            Else
                strSql = "FLOAT(" & pField.FieldSize & "," & pField.FieldDecimals & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & " NOT NULL " & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '" & pField.FieldDefault & "'"
            End If
        Else
            If pField.FieldDefault = "" Then
                strSql = "FLOAT(" & pField.FieldSize & "," & pField.FieldDecimals & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "")
            Else
                strSql = "FLOAT(" & pField.FieldSize & "," & pField.FieldDecimals & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '" & pField.FieldDefault & "'"
            End If
        End If
    End If
    
    If pField.FieldType = fieldtypeEnum Then
        If Not pField.FieldNull Then
            If pField.FieldDefault = "" Then
                strSql = "ENUM(" & pField.FieldEnumDef & ")" & " NOT NULL"
            Else
                strSql = "ENUM(" & pField.FieldEnumDef & ")" & " NOT NULL DEFAULT '" & pField.FieldDefault & "'"
            End If
        Else
            If pField.FieldDefault = "" Then
                strSql = "ENUM(" & pField.FieldEnumDef & ")"
            Else
                strSql = "ENUM(" & pField.FieldEnumDef & ")" & " DEFAULT '" & pField.FieldDefault & "'"
            End If
        End If
    End If
    
    If pField.FieldType = fieldtypeLong Then
        If Not pField.FieldNull Then
            If pField.FieldDefault = "" Then
                strSql = "BIGINT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & " NOT NULL " & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '0'"
            Else
                strSql = "BIGINT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & " NOT NULL " & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '" & pField.FieldDefault & "'"
            End If
        Else
            If pField.FieldDefault = "" Then
                strSql = "BIGINT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "")
            Else
                strSql = "BIGINT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '" & pField.FieldDefault & "'"
            End If
        End If
    End If
    
    If pField.FieldType = fieldtypeInt24 Or pField.FieldType = fieldtypeShort Then
        If Not pField.FieldNull Then
            If pField.FieldDefault = "" Then
                strSql = "INT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & " NOT NULL " & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '0'"
            Else
                strSql = "INT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & " NOT NULL " & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '" & pField.FieldDefault & "'"
            End If
        Else
            If pField.FieldDefault = "" Then
                strSql = "INT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "")
            Else
                strSql = "INT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '" & pField.FieldDefault & "'"
            End If
        End If
    End If
    
    If pField.FieldType = fieldtypeLongBlob Then
        If Not pField.FieldNull Then
            strSql = "BLOB NOT NULL"
        Else
            strSql = "BLOB"
        End If
    End If
    
    If pField.FieldType = fieldtypeLongLong Then
        If Not pField.FieldNull Then
            If pField.FieldDefault = "" Then
                strSql = "MEDIUMINT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & " NOT NULL " & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '0'"
            Else
                strSql = "MEDIUMINT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & " NOT NULL " & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '" & pField.FieldDefault & "'"
            End If
        Else
            If pField.FieldDefault = "" Then
                strSql = "MEDIUMINT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "")
            Else
                strSql = "MEDIUMINT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '" & pField.FieldDefault & "'"
            End If
        End If
    End If
    
    If pField.FieldType = fieldtypeMediumBlob Then
        If Not pField.FieldNull Then
            strSql = "MEDIUMBLOB NOT NULL"
        Else
            strSql = "MEDIUMBLOB"
        End If
    End If

    If pField.FieldType = fieldtypeSet Then
        If Not pField.FieldNull Then
            If pField.FieldDefault = "" Then
                strSql = "SET(" & pField.FieldEnumDef & ")" & " NOT NULL"
            Else
                strSql = "SET(" & pField.FieldEnumDef & ")" & " NOT NULL DEFAULT '" & pField.FieldDefault & "'"
            End If
        Else
            If pField.FieldDefault = "" Then
                strSql = "SET(" & pField.FieldEnumDef & ")"
            Else
                strSql = "SET(" & pField.FieldEnumDef & ")" & " DEFAULT '" & pField.FieldDefault & "'"
            End If
        End If
    End If
    
    If pField.FieldType = fieldtypeTime Then
        If Not pField.FieldNull Then
            If pField.FieldDefault = "" Then
                strSql = "TIME NOT NULL DEFAULT '00:00:00'"
            Else
                strSql = "TIME NOT NULL DEFAULT '" & pField.FieldDefault & "'"
            End If
        Else
            If pField.FieldDefault = "" Then
                strSql = "TIME"
            Else
                strSql = "TIME DEFAULT '" & pField.FieldDefault & "'"
            End If
        End If
    End If
    
    If pField.FieldType = fieldtypeTimeStamp Then
        strSql = "TIMESTAMP(" & pField.FieldSize & ")"
    End If
    
    If pField.FieldType = fieldtypeTiny Then
        If Not pField.FieldNull Then
            If pField.FieldDefault = "" Then
                strSql = "TINYINT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & " NOT NULL " & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '0'"
            Else
                strSql = "TINYINT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & " NOT NULL " & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '" & pField.FieldDefault & "'"
            End If
        Else
            If pField.FieldDefault = "" Then
                strSql = "TINYINT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "")
            Else
                strSql = "TINYINT(" & pField.FieldSize & ")" & IIf(pField.Unsigned, " UNSIGNED ", "") & IIf(pField.ZeroFill, " ZEROFILL ", "") & IIf(pField.FieldAutoIncrement, " AUTO INCREMENT ", "") & "DEFAULT '" & pField.FieldDefault & "'"
            End If
        End If
    End If
    
    If pField.FieldType = fieldtypeMediumBlob Then
        If Not pField.FieldNull Then
            strSql = "TINYBLOB NOT NULL"
        Else
            strSql = "TINYBLOB"
        End If
    End If
    
    If pField.FieldType = fieldtypeVarString Then
        If Not pField.FieldNull Then
            If pField.FieldDefault = "" Then
                strSql = "VARCHAR(" & pField.FieldSize & ")" & " NOT NULL"
            Else
                strSql = "VARCHAR(" & pField.FieldSize & ")" & " NOT NULL DEFAULT '" & pField.FieldDefault & "'"
            End If
        Else
            If pField.FieldDefault = "" Then
                strSql = "VARCHAR(" & pField.FieldSize & ")"
            Else
                strSql = "VARCHAR(" & pField.FieldSize & ")" & " DEFAULT '" & pField.FieldDefault & "'"
            End If
        End If
    End If
    
    If pField.FieldType = fieldtypeString Then
        If Not pField.FieldNull Then
            If pField.FieldDefault = "" Then
                strSql = "CHAR(" & pField.FieldSize & ")" & " NOT NULL"
            Else
                strSql = "CHAR(" & pField.FieldSize & ")" & " NOT NULL DEFAULT '" & pField.FieldDefault & "'"
            End If
        Else
            If pField.FieldDefault = "" Then
                strSql = "CHAR(" & pField.FieldSize & ")"
            Else
                strSql = "CHAR(" & pField.FieldSize & ")" & " DEFAULT '" & pField.FieldDefault & "'"
            End If
        End If
    End If
    
    If pField.FieldType = fieldtypeYear Then
        If Not pField.FieldNull Then
            If pField.FieldDefault = "" Then
                strSql = "YEAR(" & pField.FieldSize & ")" & " NOT NULL"
            Else
                strSql = "YEAR(" & pField.FieldSize & ")" & " NOT NULL DEFAULT '" & pField.FieldDefault & "'"
            End If
        Else
            If pField.FieldDefault = "" Then
                strSql = "YEAR(" & pField.FieldSize & ")"
            Else
                strSql = "YEAR(" & pField.FieldSize & ")" & " DEFAULT '" & pField.FieldDefault & "'"
            End If
        End If
    End If
    
    StruFieldToQuery = strSql
End Function

Function StruTable(lpTableType As enumTableType) As String
    If lpTableType = AUTO Then
        StruTable = "MyISAM"
    ElseIf lpTableType = BDB Then
        StruTable = "BDB"
    ElseIf lpTableType = HEAP Then
        StruTable = "HEAP"
    ElseIf lpTableType = INNODB Then
        StruTable = "InnoDB"
    ElseIf lpTableType = ISAM Then
        StruTable = "ISAM"
    ElseIf lpTableType = MERGE Then
        StruTable = "MERGE"
    ElseIf lpTableType = MYISAM Then
        StruTable = "MyISAM"
    End If
End Function

Function ExecuteSQL(lpSql As String, lpMySql As MYSQL, lpRecordset As MySqlRecordset, lpRetArray()) As Double
    Dim lpRecRes As MYSQL_RES
    Dim lpRecField As MYSQL_FIELD
    Dim lpRecRows As MYSQL_ROWS
    
    Dim lngRet As Long
    Dim lngFieldCount As Long
    Dim lngRowCount As Long
    Dim i As Long
    Dim j As Long
    Dim s As String
    Dim pickUp() As Long
    
    Dim pFields As MySqlFields
    
    Set lpRecordset.Fields = New MySqlFields
    Set pFields = lpRecordset.Fields
    
    lngRet = MySqlQuery(lpMySql, StrPtr(StrConv(lpSql, vbFromUnicode)))
    If lngRet = 0 Then
        lngRet = MySqlStoreResult(lpMySql)
        If lngRet Then
            CopyMemory lpRecRes, ByVal lngRet, LenB(lpRecRes)
            
            lngFieldCount = lpRecRes.field_count
            lngRowCount = Convert64ToLong(lpRecRes.row_count)
            
            lpRecordset.RecordCount = lngRowCount
            
            If lngRowCount > 0 Then
                ReDim pickUp(1 To lngFieldCount)
                ReDim lpRetArray(1 To lngRowCount, 1 To lngFieldCount)
                
                For i = 1 To lngFieldCount
                    lngRet = MySqlFetchField(lpRecRes)
                    If lngRet Then
                        CopyMemory lpRecField, ByVal lngRet, LenB(lpRecField)
                        
                        With pFields.Add(PointerToString(lpRecField.name))
                            .FieldName = PointerToString(lpRecField.name)
                            .FieldDefault = PointerToString(lpRecField.def)
                            .FieldSize = lpRecField.length
                            .FieldType = lpRecField.type
                        End With
                    End If
                Next
                
                For j = 1 To lngRowCount
                    lngRet = MySqlFetchRow(lpRecRes)
                    If lngRet Then
                        CopyMemory pickUp(1), ByVal lngRet, SIZE_OF_CHAR * lngFieldCount
                        For i = 1 To lngFieldCount
                            s = PointerToString(pickUp(i))
                            lpRetArray(j, i) = s
                        Next i
                    End If
                Next j
            End If
            
            lngRet = MySqlFreeResult(lpRecRes)
            ExecuteSQL = lngRowCount
        Else
            ExecuteSQL = Convert64ToLong(lpMySql.affected_rows)
        End If
    Else
        DefErr MySqlErrNo(lpMySql), "MySqlLib:ExecuteSQL", PointerToString(MySqlError(lpMySql))
    End If
End Function

Sub Main()
    RestoreMySqlAPI
End Sub

Private Sub RestoreMySqlAPI()
    Dim lBytes() As Byte
    Dim lngFile As Long
    
    lBytes = LoadResData("LIBMYSQL", "DLL")
    
    lngFile = FreeFile
    
    If Dir(SysDir & "libmysql.dll") = "" Then
        Open SysDir & "libmysql.dll" For Binary Shared As lngFile
        Put lngFile, , lBytes
        Close lngFile
    End If
End Sub

Public Function WinDir() As String
    Dim strTempStr As String
    Dim lngRet As Long
    
    strTempStr = Space(255)
    lngRet = GetWindowsDirectory(strTempStr, Len(strTempStr))
    WinDir = IIf(Right(left(strTempStr, lngRet), 1) = "\", left(strTempStr, lngRet), left(strTempStr, lngRet) & "\")
End Function

Public Function SysDir() As String
    Dim strTempStr As String
    Dim lngRet As Long
    
    strTempStr = Space(255)
    lngRet = GetSystemDirectory(strTempStr, Len(strTempStr))
    SysDir = IIf(Right(left(strTempStr, lngRet), 1) = "\", left(strTempStr, lngRet), left(strTempStr, lngRet) & "\")
End Function
