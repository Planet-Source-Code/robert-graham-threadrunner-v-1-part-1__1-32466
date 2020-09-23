VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CL1B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CL1B Class Definition
'
'   1st level class definition with intrinsic only property types.
'
''********************************************************************************************
Option Explicit

Private m_strL1BProperty1 As String
Private m_lngL1BProperty2 As Long

Public Property Let Property1(ByRef strProperty1 As String)
    m_strL1BProperty1 = strProperty1
End Property
Public Property Get Property1() As String
    Property1 = m_strL1BProperty1
End Property

Public Property Let Property2(ByVal lngProperty2 As Long)
    m_lngL1BProperty2 = lngProperty2
End Property
Public Property Get Property2() As Long
    Property2 = m_lngL1BProperty2
End Property

Private Function GetState() As Byte()
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "L1BProperty1", m_strL1BProperty1
        .WriteProperty "L1BProperty2", m_lngL1BProperty2
        
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
        m_strL1BProperty1 = .ReadProperty("L1BProperty1")
        m_lngL1BProperty2 = .ReadProperty("L1BProperty2")
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
