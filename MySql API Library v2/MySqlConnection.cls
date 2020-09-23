VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MySqlConnection"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_strUsername As String
Private m_strPassword As String
Private m_strServerAddress As String

Public Enum enumConnectionStatus
    statusConnected = 1
    statusDisconnected = 2
End Enum

Public Enum enumOperation
    operationCreateNew = 1
    operationCreateAlter = 8
End Enum

Public Enum enumTableType
    ISAM = 1
    MYISAM = 2
    HEAP = 3
    MERGE = 4
    INNODB = 5
    BDB = 6
    AUTO = 7
End Enum

Private m_enmStatus As enumConnectionStatus

Private m_lngConnHandle As Long
Private m_udtRec As MYSQL
Private m_udtRecRes As MYSQL_RES
Private m_udtRecField As MYSQL_FIELD
Private m_udtRecRows As MYSQL_ROWS
Private m_strServerVersion As String
Private m_pServerStatus As MySqlRecordset
Private m_pProcessLists As MySqlRecordset
Private m_pServerVariables As MySqlRecordset
Private m_strDatabaseName As String

Public Property Get ErrorNumber() As Long
    ErrorNumber = MySqlErrNo(m_udtRec)
End Property

Public Property Get ErrorDescription() As String
    ErrorDescription = PointerToString(MySqlError(m_udtRec))
End Property

Public Function DumpStructure(Optional lpTableName As String) As String
    Dim rsTmp1 As MySqlRecordset
    Dim rsTmp2 As MySqlRecordset
    Dim strDumpRet As String
    
    strDumpRet = ""
    
    If lpTableName <> "" Then
        Set rsTmp1 = Execute("SHOW CREATE TABLE `" & lpTableName & "`")
        Do While Not rsTmp1.EOF
            strDumpRet = strDumpRet & rsTmp1("Create table")
            rsTmp1.MoveNext
        Loop
    Else
        Set rsTmp2 = Execute("SHOW TABLES FROM `" & m_strDatabaseName & "`")
        Do While Not rsTmp2.EOF
            Set rsTmp1 = Execute("SHOW CREATE TABLE `" & rsTmp2("Tables_in_" & m_strDatabaseName) & "`")
            Do While Not rsTmp1.EOF
                strDumpRet = strDumpRet & rsTmp1("Create table") & ";" & vbCrLf & vbCrLf
                rsTmp1.MoveNext
            Loop
            rsTmp2.MoveNext
        Loop
    End If
    
    DumpStructure = strDumpRet
End Function

Public Function RenameField(strTableName As String, strOldName As String, strNewName As String) As Boolean
    Dim rsTmp As MySqlRecordset
    
    Set rsTmp = Execute("SHOW FIELDS FROM `" & strTableName & "`")
    Do While Not rsTmp.EOF
        If strOldName = rsTmp("field") Then
            Err.Clear
            Execute "ALTER TABLE `" & strTableName & "` CHANGE & `" & strOldName & "` `" & strNewName & "` " & rsTmp("type") & " " & IIf(rsTmp("Null") <> "YES", "NOT NULL", "") & " " & rsTmp("Extra")
            If Err Then RenameField = False Else RenameField = True
            Exit Do
        End If
        rsTmp.MoveNext
    Loop
End Function

Public Function RenameTable(strOldName As String, strNewName As String) As Boolean
    Err.Clear
    Execute "ALTER TABLE `" & strOldName & "` RENAME `" & strNewName & "`"
    If Err Then RenameTable = False Else RenameTable = True
End Function

Public Property Let DatabaseName(ByVal vData As String)
    m_strDatabaseName = vData
    UseDatabase m_strDatabaseName
End Property

Public Property Get DatabaseName() As String
    DatabaseName = m_strDatabaseName
End Property

Public Function CloseConnection() As Boolean
    If m_lngConnHandle <> 0 Then Call MySqlClose(m_udtRec)
    CloseConnection = True
End Function

Public Function DropIndex(strTableName As String, strIndexName As String) As Boolean
    If strTableName = "" Then
        DefErr vbObjectError & Len("DropIndex"), "MySqlLib:DropIndex", "Empty table name"
        DropIndex = False
    Else
        If strIndexName = "" Then
            DefErr vbObjectError & Len("DropIndex"), "MySqlLib:DropIndex", "Empty field name"
            DropIndex = False
        Else
            Err.Clear
            Execute "ALTER TABLE `" & strTableName & "` DROP INDEX `" & strIndexName & "`"
            If Err Then DropIndex = False Else DropIndex = True
        End If
    End If
End Function

Public Function DropField(strTableName As String, strFieldName As String) As Boolean
    If strTableName = "" Then
        DefErr vbObjectError & Len("DropField"), "MySqlLib:DropField", "Empty table name"
        DropField = False
    Else
        If strFieldName = "" Then
            DefErr vbObjectError & Len("DropField"), "MySqlLib:DropField", "Empty field name"
            DropField = False
        Else
            Err.Clear
            Execute "ALTER TABLE `" & strTableName & "` DROP `" & strFieldName & "`"
            If Err Then DropField = False Else DropField = True
        End If
    End If
End Function

Public Function DropTable(strTableName As String) As Boolean
    If strTableName = "" Then
        DefErr vbObjectError & Len("DropTable"), "MySqlLib:DropTable", "Empty table name"
        DropTable = False
    Else
        Err.Clear
        Execute "DROP TABLE IF EXISTS `" & strTableName & "`"
        If Err Then DropTable = False Else DropTable = True
    End If
End Function

Public Function DropDatabase(strDatabaseName As String) As Boolean
    If strDatabaseName = "" Then
        DefErr vbObjectError & Len("DropDatabase"), "MySqlLib:DropDatabase", "Empty database name"
        DropDatabase = False
    Else
        Err.Clear
        Execute "DROP DATABASE IF EXISTS `" & strDatabaseName & "`"
        If Err Then DropDatabase = False Else DropDatabase = True
    End If
End Function

Public Function CreateIndex(strTableName As String, pIndexes As MySqlIndexes) As Boolean
    Dim pIndex As MySqlIndex
    Dim pField As MySqlField
    Dim rsTmp As MySqlRecordset
    Dim strSql As String
    
    strSql = "ALTER TABLE `" & strTableName & "`, "
    
    For Each pIndex In pIndexes
        If pIndex.IndexPrimary Then
            strSql = strSql & " ADD PRIMARY KEY `" & pIndex.IndexName & "` ("
        ElseIf pIndex.IndexUnique Then
            strSql = strSql & " ADD UNIQUE `" & pIndex.IndexName & "` ("
        ElseIf pIndex.IndexFullText Then
            strSql = strSql & " ADD FULLTEXT `" & pIndex.IndexName & "` ("
        Else
            strSql = strSql & " ADD INDEX `" & pIndex.IndexName & "` ("
        End If
        
        For Each pField In pIndex.IndexFields
            strSql = strSql & "`" & pField.FieldName & "`, "
        Next
        
        strSql = RemoveLastChar(strSql, ",")
        strSql = strSql & "), "
    Next
    
    strSql = RemoveLastChar(strSql, ",")
    
    MsgBox strSql
    
    Err.Clear
    Execute strSql
    If Err Then CreateIndex = False Else CreateIndex = True
End Function

Public Function CreateField(strTableName As String, pFields As MySqlFields) As Boolean
    Dim pField As MySqlField
    Dim blnFieldExists As Boolean
    Dim rsTmp As MySqlRecordset
    Dim strSql As String
    
    strSql = "ALTER TABLE `" & strTableName & "` "
    
    Set rsTmp = Execute("SHOW FIELDS FROM `" & strTableName & "`")
    
    For Each pField In pFields
        blnFieldExists = False
        rsTmp.MoveFirst
        
        Do While Not rsTmp.EOF
            If pField.FieldName = rsTmp("field") Then
                blnFieldExists = True
                Exit Do
            End If
            rsTmp.MoveNext
        Loop
        
        If blnFieldExists Then
            strSql = strSql & " CHANGE `" & pField.FieldName & "` `" & pField.FieldName & "` " & StruFieldToQuery(pField) & ", "
        Else
            strSql = strSql & " ADD `" & pField.FieldName & "` " & StruFieldToQuery(pField) & ", "
        End If
    Next
    strSql = RemoveLastChar(strSql, ",")
    
    Err.Clear
    Execute strSql
    If Err Then CreateField = False Else CreateField = True
End Function

Public Function CreateTable(strTableName As String, pFields As MySqlFields, _
                                        pIndexes As MySqlIndexes, blnTemporary As Boolean, _
                                        lngTableType As enumTableType, lngOperation As enumOperation) As Boolean
    Dim pField As MySqlField
    Dim pIndex As MySqlIndex
    Dim pIdxField As MySqlField
    Dim blnTableExists As Boolean
    Dim blnFieldExists As Boolean
    Dim rsTmp As MySqlRecordset
    Dim strSql1 As String
    Dim strSql2 As String
    
    Set rsTmp = Execute("SHOW TABLES FROM `" & m_strDatabaseName & "`")
    Do While Not rsTmp.EOF
        If rsTmp("Tables_in_teste") = strTableName Then
            blnTableExists = True
            Exit Do
        End If
        rsTmp.MoveNext
    Loop
    
    strSql1 = "CREATE " & IIf(blnTemporary, "TEMPORARY", "") & " TABLE `" & strTableName & "` ("
    
    For Each pField In pFields
        strSql1 = strSql1 & "`" & pField.FieldName & "` " & StruFieldToQuery(pField) & ", "
    Next
    
    strSql1 = RemoveLastChar(strSql1, ",")
    
    strSql2 = ""
    For Each pIndex In pIndexes
        If pIndex.IndexPrimary Then
            strSql2 = strSql2 & " PRIMARY KEY `" & pIndex.IndexName & "` ("
        ElseIf pIndex.IndexUnique Then
            strSql2 = strSql2 & " UNIQUE `" & pIndex.IndexName & "` ("
        ElseIf pIndex.IndexFullText Then
            strSql2 = strSql2 & " FULLTEXT `" & pIndex.IndexName & "` ("
        Else
            strSql2 = strSql2 & " INDEX `" & pIndex.IndexName & "` ("
        End If
        
        For Each pIdxField In pIndex.IndexFields
            strSql2 = strSql2 & "`" & pIdxField.FieldName & "`, "
        Next
        
        strSql2 = RemoveLastChar(strSql2, ",")
        strSql2 = strSql2 & "), "
    Next
    
    strSql2 = RemoveLastChar(strSql2, ",")
    
    If strSql2 <> "" Then
        strSql1 = strSql1 & ", " & strSql2 & ")"
    Else
        strSql1 = strSql1 & ")"
    End If
    strSql1 = strSql1 & "TYPE=" & StruTable(lngTableType)
    
    If blnTableExists Then
        If lngOperation = operationCreateNew Then
            strSql1 = "DROP IF EXISTS `" & strTableName & "`;" & vbCrLf & strSql1
        Else
            strSql1 = "ALTER TABLE `" & strTableName & "` "
            
            Set rsTmp = Execute("SHOW FIELDS FROM `" & strTableName & "`")
            
            For Each pField In pFields
                blnFieldExists = False
                rsTmp.MoveFirst
                
                Do While Not rsTmp.EOF
                    If pField.FieldName = rsTmp("field") Then
                        blnFieldExists = True
                        Exit Do
                    End If
                    rsTmp.MoveNext
                Loop
                
                If blnFieldExists Then
                    strSql1 = strSql1 & " CHANGE `" & pField.FieldName & "` `" & pField.FieldName & "` " & StruFieldToQuery(pField) & ", "
                Else
                    strSql1 = strSql1 & " ADD `" & pField.FieldName & "` " & StruFieldToQuery(pField) & ", "
                End If
            Next
            strSql1 = RemoveLastChar(strSql1, ",")
            strSql1 = strSql1 & ", TYPE=" & StruTable(lngTableType)
        End If
    End If
    
    Err.Clear
    Execute strSql1
    If Err Then CreateTable = False Else CreateTable = True
End Function

Public Function CreateDatabase(strDatabaseName As String, Optional blnOverwrite As Boolean = False) As Boolean
    If strDatabaseName = "" Then
        DefErr vbObjectError & Len("CreateDatabase"), "MySqlLib:CreateDatabase", "Empty table name"
        CreateDatabase = False
    Else
        If blnOverwrite Then
            Execute "DROP DATABASE IF EXISTS `" & strDatabaseName & "`"
        End If
        
        Err.Clear
        Execute "CREATE DATABASE `" & strDatabaseName & "`"
        If Err Then CreateDatabase = False Else CreateDatabase = True
    End If
End Function

Public Function Execute(lpSql As String, Optional pRetArray) As MySqlRecordset
    Dim rsRecordset As New MySqlRecordset
    Dim pResArray()
    Dim intScriptLoopCounter As Integer
    
    Dim strScripts() As String
    
    strScripts = Split(lpSql, ";")
    
    For intScriptLoopCounter = LBound(strScripts) To UBound(strScripts)
        ExecuteSQL strScripts(intScriptLoopCounter), m_udtRec, rsRecordset, pResArray
        
        Set rsRecordset.ActiveConnection = Me
        
        rsRecordset.QueryString = strScripts(intScriptLoopCounter)
        rsRecordset.StoreResult pResArray
        
        Set Execute = rsRecordset
        
        If Not IsMissing(pRetArray) Then pRetArray = pResArray
    Next
End Function

Public Function UseDatabase(strDatabaseName As String) As Boolean
    Dim lngRet As Long
    If m_lngConnHandle <> 0 Then
        lngRet = MySqlSelectDb(m_udtRec, StrPtr(StrConv(strDatabaseName, vbFromUnicode)))
        If lngRet <> 0 Then
            DefErr MySqlErrNo(m_udtRec), "MySqlLib:UseDatabase", PointerToString(MySqlError(m_udtRec))
            UseDatabase = False
        Else
            m_strDatabaseName = strDatabaseName
            UseDatabase = True
        End If
    End If
End Function

Public Property Get ServerVariables() As MySqlRecordset
    Set m_pServerVariables = Execute("SHOW VARIABLES")
    
    Set ServerVariables = m_pServerVariables
End Property

Public Property Set ServerVariables(vData As MySqlRecordset)
    Set m_pServerVariables = vData
End Property

Public Property Get ProcessLists() As MySqlRecordset
    Set m_pProcessLists = Execute("SHOW FULL PROCESSLIST")
    
    Set ProcessLists = m_pProcessLists
End Property

Public Property Set ProcessLists(vData As MySqlRecordset)
    Set m_pProcessLists = vData
End Property

Public Property Get ServerStatus() As MySqlRecordset
    Set m_pServerStatus = Execute("SHOW STATUS")
    
    Set ServerStatus = m_pServerStatus
End Property

Public Property Set ServerStatus(vData As MySqlRecordset)
    Set m_pServerStatus = vData
End Property

Private Sub Class_Initialize()
    m_enmStatus = statusDisconnected
End Sub

Private Sub Class_Terminate()
    CloseConnection
    Set m_pServerVariables = Nothing
    Set m_pProcessLists = Nothing
    Set m_pServerStatus = Nothing
End Sub

Public Property Let ServerVersion(ByVal vData As String)
    m_strServerVersion = vData
End Property

Public Property Get ServerVersion() As String
    ServerVersion = m_strServerVersion
End Property

Public Function Connect(Optional lpUsername As String, Optional lpPassword As String, _
                                    Optional lpServerAddress As String) As Boolean
    Dim strUsername As String
    Dim strPassword As String
    Dim strServerAddress As String
    
    If lpUsername = "" Then strUsername = Username Else strUsername = lpUsername
    If lpPassword = "" Then strPassword = Password Else strPassword = lpPassword
    If lpServerAddress = "" Then strServerAddress = ServerAddress Else strServerAddress = lpServerAddress
    
    m_lngConnHandle = MySqlConnect(m_udtRec, StrPtr(StrConv(strServerAddress, vbFromUnicode)), StrPtr(StrConv(strUsername, vbFromUnicode)), StrPtr(StrConv(strPassword, vbFromUnicode)))
    
    ServerVersion = PointerToString(m_udtRec.server_version)
    
    If m_lngConnHandle Then
        CopyMemory m_udtRec, ByVal m_lngConnHandle, LenB(m_udtRec)
        Connect = True
        Username = strUsername
        Password = strPassword
        ServerAddress = strServerAddress
        Status = statusConnected
    Else
        Status = statusDisconnected
        DefErr MySqlErrNo(m_udtRec), "MySqlLib:Connect", PointerToString(MySqlError(m_udtRec))
        Connect = False
    End If
End Function

Public Property Let Status(ByVal vData As enumConnectionStatus)
    m_enmStatus = vData
End Property

Public Property Get Status() As enumConnectionStatus
    Status = m_enmStatus
End Property

Public Property Let ServerAddress(ByVal vData As String)
    m_strServerAddress = vData
End Property

Public Property Get ServerAddress() As String
    ServerAddress = m_strServerAddress
End Property

Public Property Let Password(ByVal vData As String)
    m_strPassword = vData
End Property

Public Property Get Password() As String
    Password = m_strPassword
End Property

Public Property Let Username(ByVal vData As String)
    m_strUsername = vData
End Property

Public Property Get Username() As String
    Username = m_strUsername
End Property
