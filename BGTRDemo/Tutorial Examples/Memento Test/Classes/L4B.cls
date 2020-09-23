VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CL4B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3C3CCE950136"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CL4B Class Definition
'
'   4th level class definition with only simple intrinsic property types.
'
''********************************************************************************************
Option Explicit

Private m_lngL4BProperty1 As Long
Private m_bL4BProperty2 As Boolean

Public Property Let L4BProperty1(ByVal lngL4BProperty1 As Long)
    m_lngL4BProperty1 = lngL4BProperty1
End Property
Public Property Get L4BProperty1() As Long
    L4BProperty1 = m_lngL4BProperty1
End Property

Public Property Let L4BProperty2(ByVal bL4BProperty2 As Boolean)
    m_bL4BProperty2 = bL4BProperty2
End Property
Public Property Get L4BProperty2() As Boolean
    L4BProperty2 = m_bL4BProperty2
End Property

Private Function GetState() As Byte()
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "L4BProperty1", m_lngL4BProperty1
        .WriteProperty "L4BProperty2", m_bL4BProperty2
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
        m_lngL4BProperty1 = .ReadProperty("L4BProperty1")
        m_bL4BProperty2 = .ReadProperty("L4BProperty2")
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

