VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CL4A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3C3CCE9402C6"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CL4A Class Definition
'
'   4th level class definition with only simple intrinsic property types.
'
''********************************************************************************************
Option Explicit

Private m_lngL4AProperty1 As Long
Private m_lngL4AProperty2 As Long
Private m_strL4AProperty3 As String

Public Property Let L4AProperty1(ByVal lngL4AProperty1 As Long)
    m_lngL4AProperty1 = lngL4AProperty1
End Property
Public Property Get L4AProperty1() As Long
    L4AProperty1 = m_lngL4AProperty1
End Property

Public Property Let L4AProperty2(ByVal lngL4AProperty2 As Long)
    m_lngL4AProperty2 = lngL4AProperty2
End Property
Public Property Get L4AProperty2() As Long
    L4AProperty2 = m_lngL4AProperty2
End Property

Public Property Let L4AProperty3(ByVal strL4AProperty3 As String)
    m_strL4AProperty3 = strL4AProperty3
End Property
Public Property Get L4AProperty3() As String
    L4AProperty3 = m_strL4AProperty3
End Property

Private Function GetState() As Byte()
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "L4AProperty1", m_lngL4AProperty1
        .WriteProperty "L4AProperty2", m_lngL4AProperty2
        .WriteProperty "L4AProperty3", m_strL4AProperty3
        GetState = .Contents
    End With
    
    Set pbTemp = Nothing
End Function

Public Function GetSuperState() As Byte()
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "State", GetState
        GetSuperState = .Contents
    End With
    
    Set pbTemp = Nothing

End Function
Private Sub SetState(ByRef baState() As Byte)
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = baState
        m_lngL4AProperty1 = .ReadProperty("L4AProperty1")
        m_lngL4AProperty2 = .ReadProperty("L4AProperty2")
        m_strL4AProperty3 = .ReadProperty("L4AProperty3")
    End With
    
    Set pbTemp = Nothing
End Sub

Public Sub SetSuperState(ByRef baState() As Byte)
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = baState
        SetState (.ReadProperty("State"))
    End With
    
    Set pbTemp = Nothing
End Sub
