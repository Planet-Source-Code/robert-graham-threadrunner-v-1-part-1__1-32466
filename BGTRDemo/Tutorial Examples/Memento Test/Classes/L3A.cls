VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CL3A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3C3CCE9202DA"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CL3A Class Definition
'
'   3rd level class definition with intrinsic and complex property types.
'
''********************************************************************************************
Option Explicit

'Instrinsic variables - Simple types
Private m_lngL3AProperty1 As Long
Private m_strL3AProperty2 As String

'Object variables - Complex types
Private m_CL4A As CL4A

Public Property Let L3AProperty1(ByVal lngL3AProperty1 As Long)
    m_lngL3AProperty1 = lngL3AProperty1
End Property
Public Property Get L3AProperty1() As Long
    L3AProperty1 = m_lngL3AProperty1
End Property

Public Property Let L3AProperty2(ByVal strL3AProperty2 As String)
    m_strL3AProperty2 = strL3AProperty2
End Property
Public Property Get L3AProperty2() As String
    L3AProperty2 = m_strL3AProperty2
End Property

Public Property Get L4A() As CL4A
    Set L4A = m_CL4A
End Property

Public Property Let L4A(ByVal objL4A As CL4A)
    Set m_CL4A = objL4A
End Property

Private Function GetState() As Byte()
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "L3AProperty1", m_lngL3AProperty1
        .WriteProperty "L3AProperty2", m_strL3AProperty2
        
        GetState = .Contents
    End With
    
    Set pbTemp = Nothing
End Function

Public Function GetSuperState() As Byte()
Dim pbTemp As PropertyBag
Dim i As Long


    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "State", GetState
        .WriteProperty "L4A", m_CL4A.GetSuperState
        GetSuperState = .Contents
    End With
    
    Set pbTemp = Nothing

End Function
Private Sub SetState(ByRef baState() As Byte)
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = baState
        m_lngL3AProperty1 = .ReadProperty("L3AProperty1")
        m_strL3AProperty2 = .ReadProperty("L3AProperty2")
    End With
    
    Set pbTemp = Nothing
End Sub

Public Sub SetSuperState(ByRef baState() As Byte)
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = baState
        SetState (.ReadProperty("State"))
        Set m_CL4A = New CL4A
            m_CL4A.SetSuperState (.ReadProperty("L4A"))
    End With
    
    Set pbTemp = Nothing
End Sub



