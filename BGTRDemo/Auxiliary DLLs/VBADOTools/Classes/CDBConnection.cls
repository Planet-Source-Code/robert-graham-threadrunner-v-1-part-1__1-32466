VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CDBConnection Class Definition
'   Class defines properties for saving connection information
'
'   Instancing is set to:  5 - MultiUse
''
'********************************************************************************************
Option Explicit

Private m_strServer As String
Private m_strDatabase As String
Private m_strUser As String
Private m_strPassword As String

Public Property Let Server(ByVal strServer As String)
    m_strServer = strServer
End Property
Public Property Get Server() As String
    Server = m_strServer
End Property

Public Property Let Database(ByVal strDatabase As String)
    m_strDatabase = strDatabase
End Property
Public Property Get Database() As String
    Database = m_strDatabase
End Property

Public Property Let User(ByVal strUser As String)
    m_strUser = strUser
End Property
Public Property Get User() As String
    User = m_strUser
End Property

Public Property Let Password(ByVal strPassword As String)
    m_strPassword = strPassword
End Property
Public Property Get Password() As String
    Password = m_strPassword
End Property

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Serialization/Deserialization Functions
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Function GetState() As Byte()
On Error GoTo CatchErr
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        
        .WriteProperty "Server", m_strServer
        .WriteProperty "Database", m_strDatabase
        .WriteProperty "User", m_strUser
        .WriteProperty "Password", m_strPassword
        GetState = .Contents
    End With
    Set pbTemp = Nothing

Exit Function
CatchErr:
    Err.Raise Err.Number, "VBADOTools.CDBConnection.GetState", Err.Description
End Function

Public Function GetSuperState() As Byte()
On Error GoTo CatchErr
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "State", GetState
        GetSuperState = .Contents
    End With
    Set pbTemp = Nothing

Exit Function
CatchErr:
    Err.Raise Err.Number, "VBADOTools.CDBConnection.GetSuperState", Err.Description
End Function

Private Sub SetState(ByRef arState() As Byte)
On Error GoTo CatchErr
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = arState
        m_strServer = .ReadProperty("Server")
        m_strDatabase = .ReadProperty("Database")
        m_strUser = .ReadProperty("User")
        m_strPassword = .ReadProperty("Password")
    End With
    Set pbTemp = Nothing

Exit Sub
CatchErr:
    Err.Raise Err.Number, "VBADOTools.CDBConnection.SetState", Err.Description
End Sub

Public Sub SetSuperState(ByRef arState() As Byte)
On Error GoTo CatchErr
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = arState
        SetState .ReadProperty("State")
    End With
    Set pbTemp = Nothing

Exit Sub
CatchErr:
      Err.Raise Err.Number, "VBADOTools.CDBConnection.SetSuperState", Err.Description
End Sub
