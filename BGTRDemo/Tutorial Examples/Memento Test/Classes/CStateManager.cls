VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CStateManager Class Definition
'
'   Top level manager class definition with complex only property types.
'
''********************************************************************************************
Option Explicit

Private m_oL1A As CL1A
Private m_oL1B As CL1B


Public Property Let L1A(ByRef oL1A As CL1A)
    Set m_oL1A = oL1A
End Property
Public Property Get L1A() As CL1A
    Set L1A = m_oL1A
End Property

Public Property Let L1B(ByRef oL1B As CL1B)
    Set m_oL1B = oL1B
End Property
Public Property Get L1B() As CL1B
    Set L1B = m_oL1B
End Property

Private Function GetState() As Byte()
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        GetState = .Contents
    End With
    
    Set pbTemp = Nothing
End Function

Public Function GetSuperState() As Byte()
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "State", GetState
        .WriteProperty "L1A", m_oL1A.GetSuperState
        .WriteProperty "L1B", m_oL1B.GetSuperState
        GetSuperState = .Contents
    End With
    
    Set pbTemp = Nothing

End Function
Private Sub SetState(ByRef baState() As Byte)
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = baState
    End With
    
    Set pbTemp = Nothing
End Sub

Public Sub SetSuperState(ByRef baState() As Byte)
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = baState
        SetState (.ReadProperty("State"))
        Set m_oL1A = New CL1A
        m_oL1A.SetSuperState (.ReadProperty("L1A"))
        Set m_oL1B = New CL1B
        m_oL1B.SetSuperState (.ReadProperty("L1B"))
    End With
    Set pbTemp = Nothing
End Sub
