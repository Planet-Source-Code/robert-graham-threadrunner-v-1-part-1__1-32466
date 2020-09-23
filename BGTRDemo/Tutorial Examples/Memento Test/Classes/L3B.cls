VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CL3B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3C3CCE940032"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_CL4B As CL4B
Private m_colL4Cs As CL4Cs

Public Property Get L4B() As CL4B
    Set L4B = m_CL4B
End Property
Public Property Let L4B(ByVal objL4B As CL4B)
    Set m_CL4B = objL4B
End Property

Public Property Get L4Cs() As CL4Cs
    Set L4Cs = m_colL4Cs
End Property
Public Property Let L4Cs(ByVal objL4Cs As CL4Cs)
    If Not m_colL4Cs Is Nothing Then
        Set m_colL4Cs = Nothing
    End If
    Set m_colL4Cs = objL4Cs
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
Dim i As Long
Dim oL4C As CL4C


    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "State", GetState
        .WriteProperty "L4B", m_CL4B.GetSuperState
        
        .WriteProperty "L4CCount", m_colL4Cs.Count
        i = 1
        For Each oL4C In m_colL4Cs
            .WriteProperty "L4CInstance" & i, oL4C.GetSuperState
            Set oL4C = Nothing
            i = i + 1
        Next oL4C
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
Dim i As Long
Dim Count As Long
Dim oL4C As CL4C


    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = baState
        SetState (.ReadProperty("State"))
        Set m_CL4B = New CL4B
            m_CL4B.SetSuperState (.ReadProperty("L4B"))
           
        Count = .ReadProperty("L4CCount")
        Set m_colL4Cs = New CL4Cs
        
        For i = 1 To Count
            Set oL4C = New CL4C
            oL4C.SetSuperState (.ReadProperty("L4CInstance" & i))
            m_colL4Cs.Insert oL4C
            Set oL4C = Nothing
        Next i

    End With
    
    Set pbTemp = Nothing
End Sub

Private Sub Class_Initialize()
    Set m_colL4Cs = New CL4Cs
End Sub

Private Sub Class_Terminate()
    Set m_colL4Cs = Nothing
End Sub

