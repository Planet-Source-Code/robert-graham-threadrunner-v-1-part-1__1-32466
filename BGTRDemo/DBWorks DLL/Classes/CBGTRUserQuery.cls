VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBGTRUserQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CBGTRUserQuery Class Definition
'   Class defines tasks for the ThreadRunner using a user defined query
'
'   Instancing is set to:  5 - MultiUse
''
'********************************************************************************************
Option Explicit

Implements IBGTRWork

Private m_baResults() As Byte
Private m_bIsTaskCancel As Boolean
Private m_pCancel As Long

Public Enum DBType_enum
    dbtype_Unknown = 0
    dbtype_SQLServer = 1
    dbtype_Jet = 2
End Enum

Private m_lngDBType As DBType_enum
Private m_lngQueryType As QueryType_enum
Private m_strSQLStatement As String
Private m_arParameters() As Variant
Private m_lngParameterCount As Long
Private m_lngReturnsRecords As ReturnsRecords_enum
Private m_lngUsesParameters As UsesParameters_enum

Private m_CDBConnection As CDBConnection
Private m_strConnection As String

Private m_CADOProc As CADOProc
Private WithEvents rsTarget As ADODB.Recordset
Attribute rsTarget.VB_VarHelpID = -1

Private lngFetchCount As Long


'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^  Required IBGTRWork interface implementation
'^  By making these simple wrapper functions that
'^  delegate the call, they can be used as is from
'^ class to class.
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Private Property Get IBGTRWork_WorkResults() As Byte()
    IBGTRWork_WorkResults = m_baResults
End Property

Private Sub IBGTRWork_SynchCancelPointers(ByVal pCancel As Long)
    m_pCancel = pCancel
End Sub

Private Sub IBGTRWork_InitializeWork()
    InitializeWork
End Sub

Private Sub IBGTRWork_ExecuteWork()
    ExecuteWork
End Sub

Private Sub IBGTRWork_TerminateWork()
    TerminateWork
End Sub
Private Function IBGTRWork_GetWorkState() As Byte()
    IBGTRWork_GetWorkState = GetSuperState
End Function
Private Sub IBGTRWork_SetWorkState(arWorkMemento() As Byte)
    SetSuperState arWorkMemento
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Public interface exposed to client only, not Worker thread
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public Property Let DBType(ByVal lngDBType As DBType_enum)
    m_lngDBType = lngDBType
End Property
Public Property Get DBType() As DBType_enum
    DBType = m_lngDBType
End Property

Public Property Let Connection(ByRef oCDBConnection As CDBConnection)
    Set m_CDBConnection = oCDBConnection
End Property
Public Property Get Connection() As CDBConnection
    Set Connection = m_CDBConnection
End Property

Public Property Let QueryType(ByVal lngQueryType As QueryType_enum)
    m_lngQueryType = lngQueryType
End Property
Public Property Get QueryType() As QueryType_enum
    QueryType = m_lngQueryType
End Property

Public Property Let SQLStatement(ByVal strSQLStatement As String)
    m_strSQLStatement = strSQLStatement
End Property
Public Property Get SQLStatement() As String
    SQLStatement = m_strSQLStatement
End Property

Public Property Let ADOParameters(ByRef arParameters As Variant)
    m_arParameters = arParameters
End Property
Public Property Get ADOParameters() As Variant
    ADOParameters = m_arParameters
End Property

Public Property Let ParameterCount(ByVal lngParameterCount As Long)
    m_lngParameterCount = lngParameterCount
End Property
Public Property Get ParameterCount() As Long
    ParameterCount = m_lngParameterCount
End Property

Public Property Let ReturnsRecords(ByVal lngReturnsRecords As ReturnsRecords_enum)
    m_lngReturnsRecords = lngReturnsRecords
End Property
Public Property Get ReturnsRecords() As ReturnsRecords_enum
    ReturnsRecords = m_lngReturnsRecords
End Property

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Task specific implementation functions
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub InitializeWork()
On Error GoTo CatchErr
    
    m_strConnection = MakeConnectionString
    If m_lngParameterCount > 0 Then
        m_lngUsesParameters = uprm_UsesParameters
    End If
    
Exit Sub
CatchErr:
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRUserQuery.InitializeWork", Err.Description
End Sub

Public Sub ExecuteWork()
On Error GoTo CatchErr
Dim i As Long
Dim lngRecordsAffected As Long
Dim cnConnection As ADODB.Connection
Dim tmpADOStream As ADODB.Stream
Dim tmpSize As String
Dim lngSize As Long

    'Create and open connection
    Set cnConnection = New ADODB.Connection
    cnConnection.ConnectionString = m_strConnection
    cnConnection.Open
    
    'Populate query parameters if required
    If m_lngUsesParameters = uprm_UsesParameters Then
        m_CADOProc.ParameterCount = m_lngParameterCount
        For i = 0 To m_lngParameterCount - 1
            'Size is required only for variable length data - strings generally
            tmpSize = m_arParameters(i, paramcol_Size)
            If Len(tmpSize) > 0 Then
                lngSize = CLng(tmpSize)
            Else
                lngSize = 0
            End If
            
            m_CADOProc.SetParameter CStr(m_arParameters(i, paramcol_Name)), _
                                                       CLng(m_arParameters(i, paramcol_Type)), _
                                                       CLng(m_arParameters(i, paramcol_Direction)), _
                                                       lngSize, _
                                                       m_arParameters(i, paramcol_Value)
        Next i
    End If
    
    
    Set rsTarget = New ADODB.Recordset
    
    'Execute based upon type - SQLString or StoredProc
    Select Case m_lngQueryType
    
        Case querytype_SQLString
            m_CADOProc.ExecuteCommand m_strSQLStatement, _
                                                                                            cnConnection, _
                                                                                            m_lngUsesParameters, _
                                                                                            m_lngReturnsRecords, _
                                                                                            rsTarget, _
                                                                                            rsevents_RaiseEvents
                                                                                            
        Case querytype_StoredProc
                m_CADOProc.ExecuteProcedure m_strSQLStatement, _
                                                                                            cnConnection, _
                                                                                            m_lngUsesParameters, _
                                                                                            m_lngReturnsRecords, _
                                                                                            rsTarget, _
                                                                                            rsevents_RaiseEvents
                                                                                            

    
    End Select
    
    'Check cancel before returning the results
    
    If CheckCancel Then
        rsTarget.Close
        Set rsTarget = Nothing
        Exit Sub
    End If
    
    'Serve up the recordset as a binary stream if records are returned
    If m_lngReturnsRecords = retrs_ReturnsRecords Then
        Set tmpADOStream = New ADODB.Stream

        With rsTarget
            .Save tmpADOStream, adPersistADTG
            .Close
        End With
    
        Set rsTarget = Nothing
        m_baResults = tmpADOStream.Read

    Else
        'Set the results array to default
        ReDim m_baResults(0)
        m_baResults = g_arNoData
    End If
Exit Sub
CatchErr:
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRUserQuery.ExecuteWork", Err.Description
End Sub

Private Sub TerminateWork()
    'no implementation for this task
End Sub

Private Function CheckCancel() As Boolean
    CopyMemory m_bIsTaskCancel, ByVal m_pCancel, BYTES_INT
    CheckCancel = m_bIsTaskCancel
End Function

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Various task support and auxiliary functions
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public Function MakeConnectionString() As String
On Error GoTo CatchErr
    
    Select Case m_lngDBType
    
        Case dbtype_SQLServer
        
            MakeConnectionString = _
                           "Provider=" & LoadResString(RES_SQL_PROVIDER) _
                & ";" & "Driver=" & LoadResString(RES_SQL_DRIVER) _
                & ";" & "Server=" & m_CDBConnection.Server _
                & ";" & "Database=" & m_CDBConnection.Database _
                & ";" & "UID=" & m_CDBConnection.User _
                & ";" & "PWD=" & m_CDBConnection.Password _
                & ";"
                
        Case dbtype_Jet
            
            MakeConnectionString = _
                           "Provider=" & LoadResString(RES_JET_PROVIDER) _
                & ";" & "Data Source=" & m_CDBConnection.Server _
                & ";" & "User ID=" & m_CDBConnection.User _
                & ";" & "Password=" & m_CDBConnection.Password _
                & ";"
        
    End Select
     
Exit Function
CatchErr:
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRUserQuery.MakeConnectionString", Err.Description
End Function

Private Sub rsTarget_FetchProgress(ByVal Progress As Long, ByVal MaxProgress As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    lngFetchCount = lngFetchCount + 1
    If lngFetchCount > CHECK_CANCEL_INCR Then
        lngFetchCount = 0
        If CheckCancel Then
            rsTarget.Cancel
        End If
    End If
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Serialization/Deserialization Functions
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Function GetState() As Byte()
On Error GoTo CatchErr
Dim pbTemp As PropertyBag
Dim i As Long
Dim j As Long

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "DBType", m_lngDBType
        .WriteProperty "QueryType", m_lngQueryType
        .WriteProperty "SQLStatement", m_strSQLStatement
        .WriteProperty "ParameterCount", m_lngParameterCount
        If m_lngParameterCount > 0 Then
            For i = 0 To m_lngParameterCount - 1
                For j = paramcol_Name To paramcol_Value
                    .WriteProperty "Parameters_" & i & "_" & j, m_arParameters(i, j)
                Next j
            Next i
        End If
        .WriteProperty "ReturnsRecords", m_lngReturnsRecords
        
        GetState = .Contents
    End With
    Set pbTemp = Nothing
Exit Function
CatchErr:
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRUserQuery.GetState", Err.Description
End Function

Public Function GetSuperState() As Byte()
On Error GoTo CatchErr
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
    
        .WriteProperty "State", GetState
        .WriteProperty "Connection", m_CDBConnection.GetSuperState
        GetSuperState = .Contents
        
    End With
    Set pbTemp = Nothing
Exit Function
CatchErr:
    Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRUserQuery.GetSuperState", Err.Description
End Function

Private Sub SetState(ByRef arState() As Byte)
On Error GoTo CatchErr
Dim pbTemp As PropertyBag
Dim i As Long
Dim j As Long

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = arState
        m_lngDBType = .ReadProperty("DBType")
        m_lngQueryType = .ReadProperty("QueryType")
        m_strSQLStatement = .ReadProperty("SQLStatement")
        m_lngParameterCount = .ReadProperty("ParameterCount")
                        
        If m_lngParameterCount > 0 Then
            ReDim m_arParameters(0 To m_lngParameterCount - 1, paramcol_Name To paramcol_Value)
            For i = 0 To m_lngParameterCount - 1
                For j = paramcol_Name To paramcol_Value
                    m_arParameters(i, j) = .ReadProperty("Parameters_" & i & "_" & j)
                Next j
            Next i
        End If
             
        m_lngReturnsRecords = .ReadProperty("ReturnsRecords")
        
    End With
    Set pbTemp = Nothing
Exit Sub
CatchErr:
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRUserQuery.SetState", Err.Description
End Sub

Public Sub SetSuperState(ByRef arState() As Byte)
On Error GoTo CatchErr
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = arState
        SetState .ReadProperty("State")
        
        Set m_CDBConnection = New CDBConnection
        m_CDBConnection.SetSuperState .ReadProperty("Connection")
    End With
    Set pbTemp = Nothing
Exit Sub
CatchErr:
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRUserQuery.SetSuperState", Err.Description
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Class Constructor/Destructor
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub Class_Initialize()
    Set m_CDBConnection = New CDBConnection
    Set m_CADOProc = New CADOProc
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Set m_CDBConnection = Nothing
    Set m_CADOProc = Nothing
End Sub



