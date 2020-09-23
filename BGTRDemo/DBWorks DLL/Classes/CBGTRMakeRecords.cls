VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBGTRMakeRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CBGTRMakeRecords Class Definition
'   Class defines tasks for the ThreadRunner using disconnected
'   recordset and sample Jet database
'
'   Instancing is set to:  5 - MultiUse
''
'********************************************************************************************
Option Explicit

Implements IBGTRWork

Private m_baResults() As Byte
Private m_bIsTaskCancel As Boolean
Private m_pCancel As Long


Private m_lngLastRecord As Long
Private m_dteLastDate As Date

Private m_lngRecordCount As Long
Private m_CDBConnection As VBADOTools.CDBConnection
Private m_strConnection As String

Private m_CADOProc As VBADOTools.CADOProc
Private rsTemp As ADODB.Recordset
Private WithEvents rsTarget As ADODB.Recordset
Attribute rsTarget.VB_VarHelpID = -1


Private Const START_DATE As Long = 0
Private Const END_DATE As Long = 1

'Declare hwnd for a periodic progress report in client
Private m_hwndUpdateProgress As Long

Private lngBatchCount As Long


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
Public Property Let Connection(ByRef oCDBConnection As CDBConnection)
    Set m_CDBConnection = oCDBConnection
End Property
Public Property Get Connection() As CDBConnection
    Set Connection = m_CDBConnection
End Property
'Property for number of records to create
Public Property Let FinalRecordCount(ByVal lngRecordCount As Long)
    m_lngRecordCount = lngRecordCount
End Property
Public Property Get FinalRecordCount() As Long
    FinalRecordCount = m_lngRecordCount
End Property
'Property for hwnd of a progress update window
Public Property Let UpdateHwnd(ByVal hwndUpdateProgress As Long)
    m_hwndUpdateProgress = hwndUpdateProgress
End Property
Public Property Get UpdateHwnd() As Long
    UpdateHwnd = m_hwndUpdateProgress
End Property

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Task specific implementation functions
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub InitializeWork()
On Error GoTo CatchErr

   m_strConnection = MakeConnectionString
   MakeJetProcs
   GetLastRecord
   
Exit Sub
CatchErr:
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRMakeRecords.InitializeWork", Err.Description
End Sub

Private Sub ExecuteWork()
On Error GoTo CatchErr
Dim cnConnection As ADODB.Connection
Dim dteStart As Date
Dim i As Long
Dim k As Long
Dim strMsg As String

    Set cnConnection = New ADODB.Connection
    cnConnection.ConnectionString = m_strConnection
    cnConnection.Open
       
    Set rsTarget = New ADODB.Recordset

    With rsTarget
        Set .ActiveConnection = cnConnection
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        'Set source to a query that returns no data to get table definition
        'and update capability
        .Source = "SELECT * FROM Demo1 WHERE Rec_ID < 0"
        .Open
        
        'Disconnect the recordset
       
        Set .ActiveConnection = Nothing
        cnConnection.Close
        Set cnConnection = Nothing
        
        'Populate the recordset with some simple data by decrementing
        'current date/time, and generating values from record count
        dteStart = m_dteLastDate
        For i = 1 To m_lngRecordCount
           dteStart = DateAdd("s", 1, dteStart)
          .AddNew
          .Fields(0) = i + m_lngLastRecord
          .Fields(1) = dteStart
          .Fields(2) = 1 * (i + m_lngLastRecord)
          .Fields(3) = 2 * (i + m_lngLastRecord)
          .Fields(4) = 3 * (i + m_lngLastRecord)
        
            k = k + 1
            If k = CHECK_CANCEL_INCR Then
                If m_hwndUpdateProgress > 0 Then
                    'Create a message of current count
                    strMsg = CStr(i)
                    'Update the client app UI progress textbox
                    SendMessage m_hwndUpdateProgress, WM_SETTEXT, 0, ByVal strMsg
                End If
                k = 0
                If CheckCancel Then
                    'Close the recordset
                    .Close
                    'Quit the function
                    Exit Sub
                End If
            End If
        Next i
    
        strMsg = "Update"
        SendMessage m_hwndUpdateProgress, WM_SETTEXT, 0, ByVal strMsg

        'Reset the connection
        Set cnConnection = New ADODB.Connection
        cnConnection.ConnectionString = m_strConnection
        cnConnection.Open
        k = 0
        .MoveFirst
        .ActiveConnection = cnConnection
         'Update the table in the sample database with the freshly created records
         For i = 1 To m_lngRecordCount
            .UpdateBatch adAffectCurrent
           .MoveNext
            k = k + 1
            If k = CHECK_CANCEL_INCR Then
                If m_hwndUpdateProgress > 0 Then
                    'Create a message of current count
                    strMsg = CStr(i)
                    'Update the client app UI progress textbox
                    SendMessage m_hwndUpdateProgress, WM_SETTEXT, 0, ByVal strMsg
                End If
                k = 0
                If CheckCancel Then
                    'Close the recordset
                    .Close
                    'Quit the function
                    Exit Sub
                End If
            End If
        Next i
            .Close
    End With
    cnConnection.Close
    Set cnConnection = Nothing
    Set rsTarget = Nothing
    
    
Exit Sub
CatchErr:
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRMakeRecords.ExecuteWork", Err.Description
End Sub

Private Sub TerminateWork()
On Error GoTo CatchErr
Dim strMsg As String
   'Just initialize property array - no results returned from this task
   ReDim m_baResults(0)
   m_baResults = g_arNoData
   'Send a completed message to client UI
    If m_hwndUpdateProgress > 0 Then
        strMsg = "Completed"
        SendMessage m_hwndUpdateProgress, WM_SETTEXT, 0, ByVal strMsg
   End If
   
Exit Sub
CatchErr:
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRMakeRecords.TerminateWork", Err.Description
End Sub

Private Function CheckCancel() As Boolean
    'copy yhr worker's pointer variable
    CopyMemory m_bIsTaskCancel, ByVal m_pCancel, BYTES_INT
    CheckCancel = m_bIsTaskCancel
End Function

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Various task support and auxiliary functions
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Function MakeConnectionString() As String
On Error GoTo CatchErr
                
            MakeConnectionString = _
                           "Provider=" & LoadResString(RES_JET_PROVIDER) _
                & ";" & "Data Source=" & m_CDBConnection.Server _
                & ";" & "User ID=" & m_CDBConnection.User _
                & ";" & "Password=" & m_CDBConnection.Password _
                & ";"
                
Exit Function
CatchErr:
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRMakeRecords.MakeConnectionString", Err.Description
End Function

'This function creates two simple stored proces in the sample Jet DB.
Private Sub MakeJetProcs()
On Error GoTo CatchErr
Dim oCJetProc As CJetProc
Dim cnConnection As ADODB.Connection
Dim strSQL As String
Dim arParameters() As Variant

    Set cnConnection = New ADODB.Connection
    cnConnection.ConnectionString = m_strConnection
    cnConnection.Open
    
    Set oCJetProc = New CJetProc
    With oCJetProc
        'Drop existing procs if they exist
        .DropAllJetProcedures cnConnection
        
        'Create the sp_DropDemoRecords proc - used to drop all data from the
        'the sample DB Demo1 table
        strSQL = "DELETE * FROM Demo1"
        .CreateJetProcedure cnConnection, "sp_DropDemoRecords", strSQL, 0
        
        'Create the sp_Demo_By_Date proc - used to query sample data in table Demo1
        'based upon user imputs for begin and end date ranges
        ReDim arParameters(START_DATE To END_DATE, jppcol_ParameterName To jppcol_ParameterType)
        arParameters(START_DATE, jppcol_ParameterName) = "in_StartDate"
        arParameters(START_DATE, jppcol_ParameterType) = "DATETIME"
        arParameters(END_DATE, jppcol_ParameterName) = "in_EndDate"
        arParameters(END_DATE, jppcol_ParameterType) = "DATETIME"
         
         strSQL = "SELECT * FROM Demo1 WHERE Rec_Date BETWEEN [in_StartDate] AND [in_EndDate] ORDER BY Rec_ID"
        .CreateJetProcedure cnConnection, "sp_Demo_By_Date", strSQL, 2, arParameters
    End With
    
    cnConnection.Close
    Set cnConnection = Nothing
    Set oCJetProc = Nothing

Exit Sub
CatchErr:
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRMakeRecords.MakeJetProcs", Err.Description
End Sub

'Function to check last existing records in the Demo1 sample data table
Private Sub GetLastRecord()
On Error GoTo CatchErr
Dim cnConnection As ADODB.Connection
Dim rsTemp As ADODB.Recordset
Dim strSQL As String

Dim lRetVal As Long
    
    Set cnConnection = New ADODB.Connection
    cnConnection.ConnectionString = m_strConnection
    cnConnection.Open
    
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        Set .ActiveConnection = cnConnection
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
    End With
    
    ' For some reason, this query will not work as a procedure
    strSQL = "SELECT TOP 1 * FROM Demo1 ORDER BY Rec_ID DESC"
    m_CADOProc.ExecuteCommand strSQL, _
                                                    cnConnection, _
                                                    uprm_NotUsesParameters, _
                                                    retrs_ReturnsRecords, _
                                                    rsTemp, _
                                                    rsevents_NoEvents
    'Initialize variables based upon last record, and continue adding from there
    With rsTemp
        If Not .EOF Then
            m_lngLastRecord = .Fields("Rec_ID").Value
            m_dteLastDate = .Fields("Rec_Date").Value
        Else
            m_lngLastRecord = 0
            m_dteLastDate = CDate("01/01/2000 00:00:00")
        End If
        .Close
    End With
        
    Set rsTemp = Nothing
    
    cnConnection.Close
    Set cnConnection = Nothing
    
Exit Sub
CatchErr:
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRMakeRecords.DropRecords", Err.Description
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Serialization/Deserialization Functions
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Function GetState() As Byte()
On Error GoTo CatchErr
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "RecordCount", m_lngRecordCount
        .WriteProperty "HwndProgress", m_hwndUpdateProgress
        GetState = .Contents
    End With
    Set pbTemp = Nothing
Exit Function
CatchErr:
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRMakeRecords.GetState", Err.Description
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
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRMakeRecords.GetSuperState", Err.Description
End Function

Private Sub SetState(ByRef arState() As Byte)
On Error GoTo CatchErr
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = arState
        m_lngRecordCount = .ReadProperty("RecordCount")
        m_hwndUpdateProgress = .ReadProperty("HwndProgress")
    End With
    Set pbTemp = Nothing
Exit Sub
CatchErr:
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRMakeRecords.SetState", Err.Description
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
      Err.Raise Err.Number, Err.Source & " in BGTRDBWorks.CBGTRMakeRecords.SetSuperState", Err.Description
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Class Constructor/Destructor
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub Class_Initialize()
    Set m_CDBConnection = New CDBConnection
    Set m_CADOProc = New CADOProc
    m_bIsTaskCancel = False
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Set m_CDBConnection = Nothing
    Set m_CADOProc = Nothing
End Sub



