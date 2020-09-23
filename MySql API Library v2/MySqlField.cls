VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MySqlField"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private m_strFieldName As String
Private m_enmFieldType As enumFieldTypes
Private m_lngFieldSize As Long
Private m_blnFieldNull As Boolean
Private m_strFieldDefault As String
Private m_blnFieldAutoIncrement As Boolean
Private m_varFieldValue As Variant

Public Enum enumRefreshOptions
    refreshGrant = 1
    refreshLog = 2
    refreshTables = 4
    refreshHosts = 8
    refreshStatus = 16
    refreshThreads = 32
    refreshSlave = 64
    refreshMaster = 128
    refreshReadLock = 256
    refreshFast = 32768
End Enum

Public Enum enumFieldTypes
    fieldtypeDecimal = 0
    fieldtypeTiny = 1
    fieldtypeShort = 2
    fieldtypeLong = 3
    fieldtypeFloat = 4
    fieldtypeDouble = 5
    fieldtypeNull = 6
    fieldtypeTimeStamp = 7
    fieldtypeLongLong = 8
    fieldtypeInt24 = 9
    fieldtypeDate = 10
    fieldtypeTime = 11
    fieldtypeDateTime = 12
    fieldtypeYear = 13
    fieldtypeNewDate = 14
    fieldtypeEnum = 247
    fieldtypeSet = 248
    fieldtypeTinyBlob = 249
    fieldtypeMediumBlob = 250
    fieldtypeLongBlob = 251
    fieldtypeBlob = 252
    fieldtypeVarString = 253
    fieldtypeString = 254
End Enum

Private m_blnZeroFill As Boolean
Private m_blnUnsigned As Boolean
Private m_lngFieldDecimals As Long
Private m_strFieldEnumDef As String

Public Property Let FieldEnumDef(ByVal vData As String)
    m_strFieldEnumDef = vData
End Property

Public Property Get FieldEnumDef() As String
    FieldEnumDef = m_strFieldEnumDef
End Property

Public Property Let FieldDecimals(ByVal vData As Long)
    m_lngFieldDecimals = vData
End Property

Public Property Get FieldDecimals() As Long
    FieldDecimals = m_lngFieldDecimals
End Property

Public Property Let Unsigned(ByVal vData As Boolean)
    m_blnUnsigned = vData
End Property

Public Property Get Unsigned() As Boolean
    Unsigned = m_blnUnsigned
End Property

Public Property Let ZeroFill(ByVal vData As Boolean)
    m_blnZeroFill = vData
End Property

Public Property Get ZeroFill() As Boolean
    ZeroFill = m_blnZeroFill
End Property

Public Property Let FieldValue(ByVal vData As Variant)
    m_varFieldValue = vData
End Property

Public Property Set FieldValue(ByVal vData As Variant)
    Set m_varFieldValue = vData
End Property

Public Property Get FieldValue() As Variant
Attribute FieldValue.VB_UserMemId = 0
    If IsObject(m_varFieldValue) Then
        Set FieldValue = m_varFieldValue
    Else
        FieldValue = m_varFieldValue
    End If
End Property

Public Property Let FieldAutoIncrement(ByVal vData As Boolean)
    m_blnFieldAutoIncrement = vData
End Property

Public Property Get FieldAutoIncrement() As Boolean
    FieldAutoIncrement = m_blnFieldAutoIncrement
End Property

Public Property Let FieldDefault(ByVal vData As String)
    m_strFieldDefault = vData
End Property

Public Property Get FieldDefault() As String
    FieldDefault = m_strFieldDefault
End Property

Public Property Let FieldNull(ByVal vData As Boolean)
    m_blnFieldNull = vData
End Property

Public Property Get FieldNull() As Boolean
    FieldNull = m_blnFieldNull
End Property

Public Property Let FieldSize(ByVal vData As Long)
    m_lngFieldSize = vData
End Property

Public Property Get FieldSize() As Long
    FieldSize = m_lngFieldSize
End Property

Public Property Let FieldType(ByVal vData As enumFieldTypes)
    m_enmFieldType = vData
End Property

Public Property Get FieldType() As enumFieldTypes
    FieldType = m_enmFieldType
End Property

Public Property Let FieldName(ByVal vData As String)
    m_strFieldName = vData
End Property

Public Property Get FieldName() As String
    FieldName = m_strFieldName
End Property
