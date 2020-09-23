VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MySqlFields"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"MySqlField"
Attribute VB_Ext_KEY = "Member0" ,"MySqlField"
Option Explicit

Private m_colFields As Collection

Public Function Add(Optional strKey As String) As MySqlField
    Dim objNewMember As MySqlField
    Set objNewMember = New MySqlField
    
    If Len(strKey) = 0 Then
        m_colFields.Add objNewMember
    Else
        m_colFields.Add objNewMember, strKey
    End If
    
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vIndex As Variant) As MySqlField
Attribute Item.VB_UserMemId = 0
    Set Item = m_colFields(vIndex)
End Property

Public Property Get Count() As Long
    Count = m_colFields.Count
End Property

Public Sub Remove(vIndex As Variant)
    m_colFields.Remove vIndex
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = m_colFields.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set m_colFields = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colFields = Nothing
End Sub

