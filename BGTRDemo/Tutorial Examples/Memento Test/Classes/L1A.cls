VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CL1A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3C3CCE900136"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CL1A Class Definition
'
'   1st level class definition with intrinsic and complex property types.
'
''********************************************************************************************
Option Explicit

Private m_strL1AProperty1 As String
Private m_lngL1AProperty2 As Long

Private m_colL2As As CL2As

Public Property Let Property1(ByRef strProperty1 As String)
    m_strL1AProperty1 = strProperty1
End Property
Public Property Get Property1() As String
    Property1 = m_strL1AProperty1
End Property

Public Property Let Property2(ByVal lngProperty2 As Long)
    m_lngL1AProperty2 = lngProperty2
End Property
Public Property Get Property2() As Long
    Property2 = m_lngL1AProperty2
End Property
Public Property Get L2As() As CL2As
    Set L2As = m_colL2As
End Property

Public Property Let L2As(ByVal objL2As As CL2As)
    If Not m_colL2As Is Nothing Then
        Set m_colL2As = Nothing
    End If
    Set m_colL2As = objL2As
End Property

Private Function GetState() As Byte()
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "L1AProperty1", m_strL1AProperty1
        .WriteProperty "L1AProperty2", m_lngL1AProperty2
        
        GetState = .Contents
    End With
    
    Set pbTemp = Nothing
End Function

Public Function GetSuperState() As Byte()
Dim pbTemp As PropertyBag
Dim i As Long
Dim objL2A As CL2A

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "State", GetState
        
        .WriteProperty "L2ACount", m_colL2As.Count
        i = 1
        For Each objL2A In m_colL2As
            .WriteProperty "L2AInstance" & i, objL2A.GetSuperState
            Set objL2A = Nothing
            i = i + 1
        Next objL2A

        GetSuperState = .Contents
    End With
    
    Set pbTemp = Nothing

End Function
Private Sub SetState(ByRef baState() As Byte)
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = baState
        m_strL1AProperty1 = .ReadProperty("L1AProperty1")
        m_lngL1AProperty2 = .ReadProperty("L1AProperty2")
    End With
    
    Set pbTemp = Nothing
End Sub

Public Sub SetSuperState(ByRef baState() As Byte)
Dim pbTemp As PropertyBag
Dim i As Long
Dim Count As Long
Dim objL2A As CL2A

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = baState
        SetState (.ReadProperty("State"))
        
        
        Count = .ReadProperty("L2ACount")
        Set m_colL2As = New CL2As
        
        For i = 1 To Count
            Set objL2A = New CL2A
            objL2A.SetSuperState (.ReadProperty("L2AInstance" & i))
            m_colL2As.Insert objL2A
            Set objL2A = Nothing
        Next i

    End With
    
    Set pbTemp = Nothing
End Sub

Private Sub Class_Initialize()
    Set m_colL2As = New CL2As
End Sub

Private Sub Class_Terminate()
    Set m_colL2As = Nothing
End Sub
