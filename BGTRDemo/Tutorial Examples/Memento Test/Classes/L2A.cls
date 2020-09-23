VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CL2A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3C3CCE9102B2"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CL2A Class Definition
'
'   2nd level class definition with intrinsic and complex property types.
'
''********************************************************************************************
Option Explicit

'Instrinsic variables - Simple types
Private m_lngL2AProperty1 As Long
Private m_strL2AProperty2 As String

'Object variables - Complex types
Private m_CL3A As CL3A
Private m_CL3B As CL3B

Public Property Let L2AProperty1(ByVal lngL2AProperty1 As Long)
    m_lngL2AProperty1 = lngL2AProperty1
End Property
Public Property Get L2AProperty1() As Long
    L2AProperty1 = m_lngL2AProperty1
End Property

Public Property Let L2AProperty2(ByVal strL2AProperty2 As String)
    m_strL2AProperty2 = strL2AProperty2
End Property
Public Property Get L2AProperty2() As String
    L2AProperty2 = m_strL2AProperty2
End Property

Public Property Get L3A() As CL3A
    Set L3A = m_CL3A
End Property
Public Property Let L3A(ByVal objL3A As CL3A)
    Set m_CL3A = objL3A
End Property

Public Property Get L3B() As CL3B
    Set L3B = m_CL3B
End Property
Public Property Let L3B(ByVal objL3B As CL3B)
    Set m_CL3B = objL3B
End Property

Private Function GetState() As Byte()
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "L2AProperty1", m_lngL2AProperty1
        .WriteProperty "L2AProperty2", m_strL2AProperty2
        
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
        .WriteProperty "L3A", m_CL3A.GetSuperState
        .WriteProperty "L3B", m_CL3B.GetSuperState
        GetSuperState = .Contents
    End With
    
    Set pbTemp = Nothing

End Function
Private Sub SetState(ByRef baState() As Byte)
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = baState
        m_lngL2AProperty1 = .ReadProperty("L2AProperty1")
        m_strL2AProperty2 = .ReadProperty("L2AProperty2")
    End With
    
    Set pbTemp = Nothing
End Sub

Public Sub SetSuperState(ByRef baState() As Byte)
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = baState
        SetState (.ReadProperty("State"))
        Set m_CL3A = New CL3A
            m_CL3A.SetSuperState (.ReadProperty("L3A"))
         Set m_CL3B = New CL3B
            m_CL3B.SetSuperState (.ReadProperty("L3B"))
    End With
    
    Set pbTemp = Nothing
End Sub


