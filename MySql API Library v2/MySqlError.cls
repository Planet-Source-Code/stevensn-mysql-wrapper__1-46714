VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MySqlError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_lngNumber As Long
Private m_strDescription As String
Private m_strSource As String

Public Property Let Number(ByVal vData As Long)
    m_lngNumber = vData
End Property

Public Property Get Number() As Long
    Number = m_lngNumber
End Property

Public Property Let Description(ByVal vData As String)
    m_strDescription = vData
End Property

Public Property Get Description() As String
    Description = m_strDescription
End Property

Public Property Let Source(ByVal vData As String)
    m_strSource = vData
End Property

Public Property Get Source() As String
    Source = m_strSource
End Property

Public Sub Raise(lngNumber As Long, strSource, strDescription)
    Number = lngNumber
    Source = strSource
    Description = strDescription
End Sub

Public Sub Clear()
    Number = 0
    Source = ""
    Description = ""
End Sub

Private Sub Class_Initialize()
    Clear
End Sub
