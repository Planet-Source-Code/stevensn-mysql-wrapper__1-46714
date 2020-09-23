VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MySqlIndex"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private m_strIndexName As String
Private m_blnIndexPrimary As Boolean
Private m_blnIndexUnique As Boolean
Private m_lngIndexSequence As Long
Private m_pIndexFields As MySqlFields
Private m_strIndexCollation As String
Private m_lngIndexCardinality As Long
Private m_blnIndexFullText As Boolean

Public Property Let IndexFullText(ByVal vData As Boolean)
    m_blnIndexFullText = vData
End Property

Public Property Get IndexFullText() As Boolean
    IndexFullText = m_blnIndexFullText
End Property

Public Property Let IndexCardinality(ByVal vData As Long)
    m_lngIndexCardinality = vData
End Property

Public Property Get IndexCardinality() As Long
    IndexCardinality = m_lngIndexCardinality
End Property

Public Property Let IndexCollation(ByVal vData As String)
    m_strIndexCollation = vData
End Property

Public Property Get IndexCollation() As String
    IndexCollation = m_strIndexCollation
End Property

Public Property Let IndexFields(ByVal vData As MySqlFields)
    Set m_pIndexFields = vData
End Property

Public Property Get IndexFields() As MySqlFields
    Set IndexFields = m_pIndexFields
End Property

Public Property Let IndexSequence(ByVal vData As Long)
    m_lngIndexSequence = vData
End Property

Public Property Get IndexSequence() As Long
    IndexSequence = m_lngIndexSequence
End Property

Public Property Let IndexUnique(ByVal vData As Boolean)
    m_blnIndexUnique = vData
End Property

Public Property Get IndexUnique() As Boolean
    IndexUnique = m_blnIndexUnique
End Property

Public Property Let IndexPrimary(ByVal vData As Boolean)
    m_blnIndexPrimary = vData
End Property

Public Property Get IndexPrimary() As Boolean
    IndexPrimary = m_blnIndexPrimary
End Property

Public Property Let IndexName(ByVal vData As String)
    m_strIndexName = vData
End Property

Public Property Get IndexName() As String
    IndexName = m_strIndexName
End Property
