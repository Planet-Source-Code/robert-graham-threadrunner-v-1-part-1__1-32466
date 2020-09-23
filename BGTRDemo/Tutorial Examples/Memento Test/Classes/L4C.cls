VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CL4C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3C3CCE950398"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_lngL4CProperty1 As Long

Public Property Let L4CProperty1(ByVal lngL4CProperty1 As Long)
    m_lngL4CProperty1 = lngL4CProperty1
End Property
Public Property Get L4CProperty1() As Long
    L4CProperty1 = m_lngL4CProperty1
End Property

Private Function GetState() As Byte()
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "L4CProperty1", m_lngL4CProperty1
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
        m_lngL4CProperty1 = .ReadProperty("L4CProperty1")
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

