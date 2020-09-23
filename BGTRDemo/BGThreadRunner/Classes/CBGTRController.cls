VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBGTRController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************************************
' CBGTRController Class Definition
'
' This class defines a Controller which creates Worker threads
' dispatches them to perform background tasks, all indirectly via the
' corresponding CBGTRWorkerInfo object
'
' Instancing is set to: 5 - MultiUse
'********************************************************************************************
Option Explicit

Implements ILinkTimer
Private m_CFireTimer As LinkedTimers.CFireTimer

'Declare enum for user defined errors
Public Enum ThreadRunnerErrorCodes_enum
    threadrunerr_ExceededMaxThreads = 39000
End Enum

'Declare enum for task request types
Public Enum RequestType_enum
        request_ExecuteTask = 0
        request_CancelTask = 1
End Enum

'Declare enum for task status codes
Public Enum TaskStatus_enum
        status_Completed = 0
        status_Canceled = 1
        status_Aborted = 2
        status_Pending = 3
End Enum

'Declare collections to manage the Worker pool
Private m_colWorkerPool As Collection
Private m_colActiveWorkers As Collection
Private m_colIdleWorkers As Collection

'Declare maximum number of threads to create
Private m_lngMaxWorkers As Long

'Declare local array to cache MemoryMap info
Private m_arMemoryMaps() As String
'Declare MemoryMapManager
Private m_CMemoryMapManager As SharedMemMaps32.CMemoryMapManager
'Declare variable for client reserved memory map
Private m_strClientMemoryMap As String
'Declare collections for managing memory maps
Private m_colFreeMemory As Collection
Private m_colAssignedMemory As Collection

Private m_hwndMessage As Long

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
''^ 'The following functions implement the procedures that
''^ allow the controller to dispatch itself upon notification.
''^ In order to prevent blocking the client application
''^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public Sub InitializeController(ByRef arMemoryMaps() As String, ByVal hwndMessage As Long)
On Error GoTo CatchErr
Dim i As Long
Dim Low As Long
Dim High As Long

'Cache the message handle
m_hwndMessage = hwndMessage

'Cache the memorymaps array to pass to threads
m_arMemoryMaps = arMemoryMaps

'Initialize the memory
Set m_CMemoryMapManager = New SharedMemMaps32.CMemoryMapManager
m_CMemoryMapManager.MakeMapFiles arMemoryMaps

'Just in case someone passes us a non-zero base array
Low = LBound(arMemoryMaps)
High = UBound(arMemoryMaps)

'Cache the client reserved memory map key
m_strClientMemoryMap = arMemoryMaps(Low, mapcol_MapKey)

'Populate the MemoryKeys collection
With m_colFreeMemory
    For i = Low + 1 To High
        .Add arMemoryMaps(i, mapcol_MapKey), arMemoryMaps(i, mapcol_MapKey)
    Next i
    'Set the max workers equal to the final data memory maps count
    m_lngMaxWorkers = .Count
End With

Exit Sub
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRController.InitializeController", Err.Description)
    Set m_CMemoryMapManager = Nothing
    RaiseError strErrorID
End Sub

Public Sub ExecuteTask(ByVal lngType As RequestType_enum, ByVal strTaskID As String)
On Error GoTo CatchErr
Dim strTask As String
    'Prefix curTaskID with the request code
    strTask = Str$(lngType) & strTaskID
    Call m_CFireTimer.SetNewTimer(Me, strTask, ZERO)
Exit Sub
CatchErr:
    'Should not occur unless no timers are availble
    Err.Raise Err.Number, Err.Source & " in CBGTRController.ExecuteTask", Err.Description
End Sub

Public Sub ReleaseMemory(ByVal strMemoryKey As String)
On Error GoTo CatchErr
    m_colAssignedMemory.Remove (strMemoryKey)
    m_colFreeMemory.Add strMemoryKey, strMemoryKey
Exit Sub
CatchErr:
'Should not occur - no cleanup possible because this is a collection error
    Err.Raise Err.Number, Err.Source & " in CBGTRController.ReleaseMemory", Err.Description
End Sub

Public Sub Shutdown()
On Error Resume Next
Dim i As Long
Dim tmpWorkerInfo As CBGTRWorkerInfo
    
    With m_colWorkerPool
        For i = 1 To .Count
            Set tmpWorkerInfo = .Item(FIRST_ITEM)
            tmpWorkerInfo.DoCancelTask
            .Remove (FIRST_ITEM)
            Set tmpWorkerInfo = Nothing
        Next i
    End With
End Sub

Public Sub NotifyController(ByVal strTaskID As String, ByVal strMemoryKey As String, ByVal lngStatus As TaskStatus_enum, Optional ByVal strMsg As String = vbNullChar)
On Error GoTo CatchErr
    
    UpdateTask strTaskID, strMemoryKey, lngStatus
    
Exit Sub
CatchErr:
'Should not occur.  Not much to do here
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRController.GetNexWorker", Err.Description)
    ReturnTaskError strErrorID, m_strClientMemoryMap, strTaskID
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ The following functions dispatch Worker objects to perform
'^  background tasks
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub DoExecuteTask(ByRef strTaskID As String)
On Error GoTo CatchErr
Dim tmpWorkerInfo As CBGTRWorkerInfo
Dim pWorker As Long
Dim strMemoryKey As String

    'Get/Create next worker
    pWorker = GetNextWorker
    
    'Get next memory key
    strMemoryKey = GetNextMemory
    'Copy the task memento to the memory map
    m_CMemoryMapManager.CopyMapMemory strMemoryKey, m_strClientMemoryMap, True
    
    'Retrieve the actual Worker from the Pool collection
    Set tmpWorkerInfo = m_colWorkerPool.Item(CStr(pWorker))
    'Dispatch the worker
    tmpWorkerInfo.DispatchWorker strTaskID, strMemoryKey
    'Destroy the tmpWorkerInfo object
    Set tmpWorkerInfo = Nothing
    
    'Made it this far without erroring out, so balance out the collections
    m_colActiveWorkers.Add pWorker, strTaskID
    m_colIdleWorkers.Remove (CStr(pWorker))
    m_colFreeMemory.Remove strMemoryKey
    m_colAssignedMemory.Add strMemoryKey, strMemoryKey
    '
Exit Sub
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRController.InitializeController", Err.Description)
    On Error Resume Next
    If Len(strMemoryKey) > 0 Then
        'Return error in the assigned memory map file
        'Make sure we cover all bases for the memory!
        m_colFreeMemory.Remove strMemoryKey
        'Just in case!  Cannot really ever need this because of function sequence
        m_colAssignedMemory.Remove strMemoryKey
        'Now make sure memory is assigned
        m_colAssignedMemory.Add strMemoryKey, strMemoryKey
    
        If pWorker <> 0 Then
            m_colActiveWorkers.Remove (strTaskID)
            m_colIdleWorkers.Add pWorker, (CStr(pWorker))
        End If
    Else
        'Return the error in submission map file
        strMemoryKey = m_strClientMemoryMap
    End If
    ReturnTaskError strErrorID, strMemoryKey, strTaskID
End Sub

Private Sub DoCancelTask(ByRef strTaskID As String)
On Error GoTo CatchErr
Dim tmpWorkerInfo As CBGTRWorkerInfo
Dim pWorker As Long

    'Retrieve the Worker pointer from the active collection
    pWorker = m_colActiveWorkers.Item(strTaskID)
        
   'Retrieve the actual Worker from the Pool collection and cancel task
    Set tmpWorkerInfo = m_colWorkerPool.Item(CStr(pWorker))
    With tmpWorkerInfo
        .DoCancelTask
    
        'Cancel implicitly releases the assigned memory
        m_colFreeMemory.Add .MemoryKey, .MemoryKey
        m_colAssignedMemory.Remove .MemoryKey
    End With
    
    'Destroy the temporary WorkerInfo object
    Set tmpWorkerInfo = Nothing
    
Exit Sub
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRController.DoCancelTask", Err.Description)
    On Error Resume Next
    'Make sure the memory gets cleaned up
    If Len(tmpWorkerInfo) > 0 Then
        With tmpWorkerInfo
            m_colFreeMemory.Add .MemoryKey, .MemoryKey
            m_colAssignedMemory.Remove .MemoryKey
        End With
    End If
    ReturnTaskError strErrorID, m_strClientMemoryMap, strTaskID
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ The following functions create and initialize Worker
'^ objects
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Function GetNextWorker() As Long
On Error GoTo CatchErr
Dim pWorker As Long

    'Look first for an idle worker
    With m_colIdleWorkers
        If .Count > 0 Then
            pWorker = .Item(FIRST_ITEM)
        Else
            'See if we are maxed out (should not happen!)
            If m_colWorkerPool.Count < m_lngMaxWorkers Then
                'Create a new worker
                pWorker = CreateNewWorker
                'Add the new worker to the Idle collection
                .Add pWorker, CStr(pWorker)
            Else
                'Raise an error!
                Err.Raise vbObjectError + threadrunerr_ExceededMaxThreads, vbNullString, LoadResString(threadrunerr_ExceededMaxThreads)
            End If
        End If
    End With
    GetNextWorker = pWorker
Exit Function
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRController.GetNextWorker", Err.Description)
    On Error Resume Next
    'Make sure the worker is accounted for
    If pWorker <> 0 Then
        m_colIdleWorkers.Add pWorker, CStr(pWorker)
    End If
    ReturnTaskError strErrorID, m_strClientMemoryMap
End Function

Private Function CreateNewWorker() As Long
On Error GoTo CatchErr
Dim tmpWorkerInfo As CBGTRWorkerInfo
Dim pWorker As Long

'Create a new worker thread
Set tmpWorkerInfo = New CBGTRWorkerInfo
pWorker = tmpWorkerInfo.CreateWorker(Me, m_arMemoryMaps)

'Add the Worker to the Pool, using the ObjPtr as a key
m_colWorkerPool.Add tmpWorkerInfo, CStr(pWorker)
CreateNewWorker = pWorker

Exit Function
CatchErr:
Dim strErrorID As String
    'No cleanup - errors here mean we never got a worker
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRController.CreateNewWorker", Err.Description)
    ReturnTaskError strErrorID, m_strClientMemoryMap
End Function

Private Function GetNextMemory() As String
On Error GoTo CatchErr
    GetNextMemory = m_colFreeMemory.Item(FIRST_ITEM)
Exit Function
CatchErr:
Dim strErrorID As String
    '??? Should not occur
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRController.GetNextMemory", Err.Description)
    ReturnTaskError strErrorID, m_strClientMemoryMap
End Function

Private Sub UpdateTask(ByRef strTaskID As String, ByRef strMemoryKey As String, ByVal lngStatus As TaskStatus_enum)
On Error GoTo CatchErr
    Select Case lngStatus
    
        Case status_Completed, _
                status_Canceled, _
                status_Aborted
                
                    TerminateTask strTaskID, strMemoryKey, lngStatus
                    
        Case status_Pending
                    'TO DO ????
        
    End Select
Exit Sub
CatchErr:
Dim strErrorID As String
    '??? Should not occur
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRController.UpdateTask", Err.Description)
    ReturnTaskError strErrorID, strMemoryKey, strTaskID
End Sub

Private Sub TerminateTask(ByRef strTaskID As String, ByRef strMemoryKey As String, ByVal lngStatus As TaskStatus_enum)
On Error GoTo CatchErr
Dim pWorker As Long

    'Remove the worker from the Active collection
    With m_colActiveWorkers
        pWorker = .Item(strTaskID)
        .Remove (strTaskID)
    End With
    'Return the worker to the idle collection
    m_colIdleWorkers.Add pWorker, CStr(pWorker)
    
    'Return results to client app
    NotifyClient strTaskID, strMemoryKey, lngStatus
    
Exit Sub
CatchErr:
Dim strErrorID As String
    '??? Should not occur - cannot cleanup anything because these are already collection errors
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRController.TerminateTask", Err.Description)
    ReturnTaskError strErrorID, strMemoryKey, strTaskID
End Sub

''^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ The following function provides message back to the client
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Private Sub NotifyClient(ByRef strTaskID As String, ByRef strMemoryKey As String, ByVal lngStatus As TaskStatus_enum)
On Error GoTo CatchErr
Dim strMsg As String
    
    strMsg = strTaskID & "|" & strMemoryKey & "|" & CStr(lngStatus)
    SendMessage m_hwndMessage, WM_SETTEXT, 0, ByVal strMsg
    
Exit Sub
CatchErr:
    'Should not occur, but if so, we'll have to raise a dialog here
    MsgBox Err.Number & "  " & Err.Source & " in CBGTRController.NotifyClient" & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub ReturnTaskError(ByRef strErrorID As String, ByRef strMemoryKey As String, Optional ByRef strTaskID As String = vbNullString)
On Error GoTo CatchErr
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "ErrorInfo", ReturnError(strErrorID)
        m_CMemoryMapManager.WriteMapMemory strMemoryKey, .Contents
    End With
    Set pbTemp = Nothing
    
    'Generate a unique identifier if there is no task ID
    If Len(strTaskID) = 0 Then
        strTaskID = CStr(CCur(Timer))
        strTaskID = "E" & Right(strTaskID, (Len(strTaskID) - 1))
    End If
    
    NotifyClient strTaskID, strMemoryKey, status_Aborted

Exit Sub
CatchErr:
'Not much we can do at this point!
    MsgBox "The Controller could not return error info" _
                & vbCrLf & "The task and system status is unknown!", vbCritical
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Timer implementations
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub ILinkTimer_FireTimer(ByVal strTimerLink As String, ByVal lngTimerID As Long)
    DoTimerLink strTimerLink, lngTimerID
End Sub

Private Sub DoTimerLink(ByRef strTimerLink As String, ByVal lngTimerID As Long)
On Error GoTo CatchErr
Dim lngRequestType As Long

    'Retrieve the request type from the string
    strTimerLink = Trim$(strTimerLink)
    
    lngRequestType = CLng(Left(strTimerLink, 1))
    
    'Reinstate the original TaskID
    strTimerLink = Right(strTimerLink, Len(strTimerLink) - 1)
    'Interpret the request type
    Select Case lngRequestType

        Case request_ExecuteTask
            DoExecuteTask strTimerLink
            
        Case request_CancelTask
            DoCancelTask strTimerLink
        
        Case Else
            'Raise an Error
            Err.Raise 5 'Invalid Procedure call
    End Select
Exit Sub
CatchErr:
Dim strErrorID As String
    '??? Should not occur - cannot cleanup anything because these are already collection errors
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRController.DoTimerLink", Err.Description)
    ReturnTaskError strErrorID, m_strClientMemoryMap
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Class constructor/destructor
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub Class_Initialize()
Set m_CFireTimer = New LinkedTimers.CFireTimer
Set m_colWorkerPool = New Collection
Set m_colActiveWorkers = New Collection
Set m_colIdleWorkers = New Collection
Set m_colFreeMemory = New Collection
Set m_colAssignedMemory = New Collection
End Sub

Private Sub Class_Terminate()
On Error Resume Next
Set m_CFireTimer = Nothing
Set m_colWorkerPool = Nothing
Set m_colIdleWorkers = Nothing
Set m_colActiveWorkers = Nothing
Set m_CMemoryMapManager = Nothing
Set m_colFreeMemory = Nothing
Set m_colAssignedMemory = Nothing
MsgBox "Controller is terminating!"
End Sub


