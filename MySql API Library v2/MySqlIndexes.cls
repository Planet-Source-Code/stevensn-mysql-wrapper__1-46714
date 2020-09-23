VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MySqlIndexes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MySqlIndex"
Attribute VB_Ext_KEY = "Member0" ,"MySqlIndex"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_colIndexes As Collection

Public Function Add(strKey As String, Optional sKey As String) As MySqlIndex
    Dim objNewMember As MySqlIndex
    Set objNewMember = New MySqlIndex

    If Len(sKey) = 0 Then
        m_colIndexes.Add objNewMember
    Else
        m_colIndexes.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vIndex As Variant) As MySqlIndex
Attribute Item.VB_UserMemId = 0
  Set Item = m_colIndexes(vIndex)
End Property

Public Property Get Count() As Long
    Count = m_colIndexes.Count
End Property

Public Sub Remove(vIndex As Variant)
    m_colIndexes.Remove vIndex
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colIndexes.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set m_colIndexes = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colIndexes = Nothing
End Sub
