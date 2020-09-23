VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MySqlRecordset"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MySqlFields"
Option Explicit

Private m_blnEOF As Boolean
Private m_blnBOF As Boolean
Private m_lngRowID As Long
Private m_sngPercentPosition As Single
Private m_lngNumRows As Long
Private m_pFields As MySqlFields
Private m_pActiveConnection As MySqlConnection
Private m_enmStatus As enumRecordsetStatus

Public Enum enumRecordsetStatus
    statusOpen
    statusClosed
End Enum

Dim lpResultArray()
Private m_strQueryString As String

Public Function OpenRecordset(Optional lpSql As String, Optional lpCon As MySqlConnection) As Boolean
    Dim lpRs As New MySqlRecordset
    Dim strSql As String
    Dim pConn As MySqlConnection
    
    If m_enmStatus = statusOpen Then
        DefErr vbObjectError & Len("OpenRecordset"), "MySqlLib:OpenRecordset", "Operation not allowed if the object is already open"
        OpenRecordset = False
    Else
        If lpSql = "" Then strSql = QueryString Else strSql = lpSql
        If lpCon Is Nothing Then Set pConn = ActiveConnection Else Set pConn = lpCon
        
        Set ActiveConnection = pConn
        Set lpRs = ActiveConnection.Execute(strSql, lpResultArray)
        
        Set Fields = lpRs.Fields
        RecordCount = lpRs.RecordCount
        RowID = lpRs.RowID
        
        Set lpRs = Nothing
        OpenRecordset = True
    End If
End Function

Public Function CloseRecordset()
    m_enmStatus = statusClosed
    Set m_pFields = New MySqlFields
    
    m_lngNumRows = 0
    m_lngRowID = 0
    m_blnBOF = True
    m_blnEOF = True
    m_sngPercentPosition = 0
End Function

Private Sub UpdateValues()
    Dim lngField As Long
    
    For lngField = 1 To m_pFields.Count
        m_pFields(lngField).FieldValue = lpResultArray(m_lngRowID, lngField)
    Next
End Sub

Public Property Let QueryString(ByVal vData As String)
    m_strQueryString = vData
End Property

Public Property Get QueryString() As String
    QueryString = m_strQueryString
End Property

Public Property Set ActiveConnection(ByVal vData As MySqlConnection)
    Set m_pActiveConnection = vData
End Property

Public Property Get ActiveConnection() As MySqlConnection
    Set ActiveConnection = m_pActiveConnection
End Property

Public Property Set Fields(ByVal vData As MySqlFields)
    Set m_pFields = vData
End Property

Public Property Get Fields() As MySqlFields
Attribute Fields.VB_UserMemId = 0
    If m_pFields Is Nothing Then
        Set m_pFields = New MySqlFields
    End If
    
    Set Fields = m_pFields
End Property

Sub StoreResult(pResArray())
    Dim nRows
    
    lpResultArray = pResArray

On Error Resume Next
    nRows = UBound(lpResultArray)
    
On Error GoTo 0
    RecordCount = nRows
    RowID = 1
End Sub

Public Property Let RecordCount(ByVal vData As Long)
    m_lngNumRows = vData
End Property

Public Property Get RecordCount() As Long
    RecordCount = m_lngNumRows
End Property

Public Property Let PercentPosition(ByVal vData As Single)
    m_sngPercentPosition = vData
End Property

Public Property Get PercentPosition() As Single
    PercentPosition = m_sngPercentPosition
End Property

Public Property Let RowID(ByVal vData As Long)
    m_lngRowID = vData
    
    If RowID > RecordCount Then
        PercentPosition = 100
    Else
        PercentPosition = (RowID / RecordCount) * 100
    End If
    
    Select Case RowID
        Case RecordCount + 1
            EOF = True
            BOF = False
        Case Is > RecordCount + 1
            EOF = True
            BOF = False
            DefErr vbObjectError & Len("RowID"), "MySqlLib:RowID", "Try to past end of recordset"
        Case 1
            BOF = True
            EOF = False
            UpdateValues
        Case Is < 1
            BOF = True
            EOF = False
            DefErr vbObjectError & Len("RowID"), "MySqlLib:RowID", "Try to past start of recordset"
        Case Else
            BOF = False
            EOF = False
            UpdateValues
    End Select
End Property

Public Property Get RowID() As Long
    RowID = m_lngRowID
End Property

Public Property Let BOF(ByVal vData As Boolean)
    m_blnBOF = vData
End Property

Public Property Get BOF() As Boolean
    BOF = m_blnBOF
End Property

Public Property Let EOF(ByVal vData As Boolean)
    m_blnEOF = vData
End Property

Public Property Get EOF() As Boolean
    EOF = m_blnEOF
End Property

Public Sub Refresh()
    Dim pResArray()
    Dim rsTmp As New MySqlRecordset
    Dim pMyRec As MYSQL
    
    ActiveConnection.Execute m_strQueryString, pResArray
    StoreResult pResArray
End Sub

Public Sub Move(lngRowID As Long)
    If lngRowID < 1 Or lngRowID > RecordCount Then
        DefErr vbObjectError & Len("Move"), "MySqlLib:Move", "Try to move out os recordset size"
    Else
        RowID = lngRowID
    End If
End Sub

Public Sub MoveFirst()
    RowID = 1
End Sub

Public Sub MovePrevious()
    RowID = RowID - 1
End Sub

Public Sub MoveNext()
    RowID = RowID + 1
End Sub

Public Sub MoveLast()
    RowID = RecordCount
End Sub

Private Sub Class_Initialize()
    Set m_pFields = New MySqlFields
    m_enmStatus = statusClosed
End Sub

Private Sub Class_Terminate()
    CloseRecordset
End Sub
