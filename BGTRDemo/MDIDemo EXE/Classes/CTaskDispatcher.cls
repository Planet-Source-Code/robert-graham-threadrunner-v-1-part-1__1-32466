VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTaskDispatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************************
' CTaskDispatcher Class Definition
'
' This class defines a Dispatcher which provides central management of
'  the entire process of dispatching tasks to the Controller and
'  returning results to the submittor
'
'********************************************************************************************
Option Explicit

Public Enum SubmittedTaskStatus_enum
    submitstatus_Active = 0
    submitstatus_Queued = 1
    submitstatus_Terminated = 2
    submitstatus_Canceling = 3
End Enum

Private Enum TaskInfo_enum
    taskinfo_TaskID = 0
    taskinfo_MemoryKey = 1
    taskinfo_TaskStatus = 2
End Enum

'Declare number of map files
Private m_lngMapFileCount As Long

'Declare client - Controller memory map
Private m_strTaskSubmissionMap As String

'Declare array, collection, and DLL for memory maps
Private m_arMemoryMaps() As String
Private m_colMemoryMaps As Collection
Private m_CMemoryMapManager As SharedMemMaps32.CMemoryMapManager

'Declare the Controller
Private m_CBGTRController As BGThreadRunner.CBGTRController

'Declare collections to manage task submissions
Private m_colActiveTaskRecipients As Collection
Private m_colQueuedTasks As Collection
Private m_colQueuedTaskRecipients As Collection
Private m_colQueuedTaskIDs As Collection

'Declare varaible to keep a running count of available task slots
Private m_lngAvailableTasks As Long

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Memory and ThreadRunner Initialization Functions
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public Sub InitializeTaskService()
On Error GoTo CatchErr
    InitializeThreadRunner
    Load frmTaskViewer
Exit Sub
CatchErr:
    Err.Raise Err.Number, Err.Source & " in CTaskDispatcher.InitializeTaskService", Err.Description
End Sub

Private Sub InitializeThreadRunner()
On Error GoTo CatchErr
Dim hwnd_ThreadTextWindow As Long
Dim arMemoryMaps() As String

    'Find the hidden window for the thread
    hwnd_ThreadTextWindow = FindThreadTextWindow
    'Create the MemoryMaps array and initialize client memory
    CreateMemoryMaps
    'Initialize the threadrunner component
    Set m_CBGTRController = New BGThreadRunner.CBGTRController
    m_CBGTRController.InitializeController m_arMemoryMaps, hwnd_ThreadTextWindow
    'All tasks are initially available
    m_lngAvailableTasks = m_lngMapFileCount
    
Exit Sub
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CTaskDispatcher.InitializeThreadRunner", Err.Description)
    Set m_CBGTRController = Nothing
    RaiseError strErrorID
End Sub

Private Sub CreateMemoryMaps()
On Error GoTo CatchErr
Dim strMemName As String
Dim i As Long

m_lngMapFileCount = CLng(LoadResString(RES_MAP_NUM))
strMemName = LoadResString(RES_MEM_NAME) & CStr(App.hInstance)

ReDim m_arMemoryMaps(0 To m_lngMapFileCount, mapcol_MapName To mapcol_MapSize)
    For i = 0 To m_lngMapFileCount
        m_arMemoryMaps(i, mapcol_MapKey) = strMemName & i
        m_arMemoryMaps(i, mapcol_MapName) = m_arMemoryMaps(i, mapcol_MapKey)
        m_arMemoryMaps(i, mapcol_MapSize) = CLng(LoadResString(RES_MEM_START + i))
    Next i
    
    Set m_CMemoryMapManager = New SharedMemMaps32.CMemoryMapManager
    m_CMemoryMapManager.MakeMapFiles m_arMemoryMaps

    For i = 1 To m_lngMapFileCount
        m_colMemoryMaps.Add m_arMemoryMaps(i, mapcol_MapName), m_arMemoryMaps(i, mapcol_MapKey)
    Next i

    m_strTaskSubmissionMap = m_arMemoryMaps(0, mapcol_MapKey)
    
Exit Sub
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CTaskDispatcher.CreateMemoryMaps", Err.Description)
    Set m_CMemoryMapManager = Nothing
    Set m_colMemoryMaps = Nothing
    RaiseError strErrorID
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Public Task Submission and Cancel Functions
'^ called by objects in the application to submit/cancel tasks
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public Function SubmitTask(ByRef arTaskMemento() As Byte, ByRef oReturnRecipient As Object, ByRef strDescription As String) As String
On Error GoTo CatchErr
Dim strTaskID As String
Dim arTemp() As Byte

    'Create a unique taskID
    strTaskID = MakeTaskID
    
    'All tasks get queued by default
    m_colQueuedTasks.Add arTaskMemento, strTaskID
    m_colQueuedTaskRecipients.Add oReturnRecipient, strTaskID
    m_colQueuedTaskIDs.Add strTaskID, strTaskID
    
    'Post new task info to the queue viewer form
    frmTaskViewer.UpdateTask strTaskID, strDescription, submitstatus_Queued
    
    'Determine task availability
    If m_lngAvailableTasks > 0 Then
        'A task is available - make sure the memory map file is cleared
        ReDim arTemp(0)
        arTemp = m_CMemoryMapManager.ReadMapMemory(m_strTaskSubmissionMap)
        If UBound(arTemp) > 0 Then
            'Something is still in the task submission memory
            'Wait for controller to catch up, but don't report
            'no action required here -  the task is already queued
        Else
            'The task submission memory is clear - we can submit the task
            SubmitQueue
        End If
    Else
        'No tasks are available - report this to the user
        MsgBox "All worker threads are currently active.  A maximum of " & m_lngMapFileCount & " background tasks may be run concurrently." _
        & vbCrLf & "The task submitted has been queued, pending completion of a currently running task." _
        & vbCrLf & "Queued tasks will be executed automatically, in order of submission, as resources become available." _
        & vbCrLf & "No action is required.  Active tasks and the queue may be viewed, and in some cases canceled, from the task viewer function.", vbInformation
    End If
    
    'Return the taskId to the submitter
    SubmitTask = strTaskID
    
Exit Function
CatchErr:
    Err.Raise Err.Number, Err.Source & " in CTaskDispatcher.SubmitTask", Err.Description
End Function

Public Sub CancelTask(ByRef strTaskID As String)
On Error GoTo CatchErr
    m_CBGTRController.ExecuteTask request_CancelTask, strTaskID
Exit Sub
CatchErr:
    Err.Raise Err.Number, Err.Source & " in CTaskDispatcher.CancelTask", Err.Description
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Private Task Submission support Functions
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Function MakeTaskID() As String
On Error GoTo CatchErr
    'Return current timer value as a string
    MakeTaskID = CStr(CCur(Timer))
Exit Function
CatchErr:
    Err.Raise Err.Number, Err.Source & " in CTaskDispatcher.MakeTaskID", Err.Description
End Function

Private Sub SubmitQueue()
On Error GoTo CatchErr
Dim strTaskID As String
    
    'Check task availability
    If m_lngAvailableTasks > 0 Then
        'Check the queue, and if not empty, submit the next task
        If m_colQueuedTasks.count > 0 Then
            'Retrieve the first task memento from task queue and copy to memory
            m_CMemoryMapManager.WriteMapMemory m_strTaskSubmissionMap _
                                                                             , m_colQueuedTasks.Item(FIRST_ITEM)
            'Retrieve the associated taskID from the taskID queue
            strTaskID = m_colQueuedTaskIDs.Item(FIRST_ITEM)
            'Submit to Controller
            m_CBGTRController.ExecuteTask request_ExecuteTask, strTaskID
            'Move the results recipient from queue collection to the active task collection
            m_colActiveTaskRecipients.Add m_colQueuedTaskRecipients.Item(FIRST_ITEM), strTaskID
            'Update the available task count
            m_lngAvailableTasks = m_lngAvailableTasks - 1
            'Update the queue viewer form to indicate task is now active
            frmTaskViewer.UpdateTask strTaskID, vbNullString, submitstatus_Active
            'All done - clean up the queue collections
            m_colQueuedTaskIDs.Remove (FIRST_ITEM)
            m_colQueuedTasks.Remove (FIRST_ITEM)
            m_colQueuedTaskRecipients.Remove (FIRST_ITEM)
        End If
    End If
    
Exit Sub
CatchErr:
    MsgBox Err.Description & Err.Source & " in CTaskDispatcher.SubmitQueue" & Err.Number, vbCritical
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Public Task Update Function called by MDI main form
'^ after receipt of a msg in hidden window
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public Sub NotifyTaskUpdate(ByRef strMsg As String)
On Error GoTo CatchErr
Dim arTaskInfo() As String
Dim strTaskID As String

Dim lngStatus As Long
Dim strMemoryKey As String

    'Parse the message string
    arTaskInfo = Split(strMsg, "|")
    
    'Assign elements to variables
    strTaskID = arTaskInfo(taskinfo_TaskID)
    strMemoryKey = arTaskInfo(taskinfo_MemoryKey)
    lngStatus = CLng(arTaskInfo(taskinfo_TaskStatus))
    'Check for Controller errors
    If Left(strTaskID, 1) = "E" Then
        'Process the error
        ProcessControllerError strMemoryKey
    Else
        'Call function to notify recipient
        DeliverResults strTaskID, lngStatus, strMemoryKey
    End If
    
    ' Do the queue - at least one task is now available
    SubmitQueue
Exit Sub
CatchErr:
    MsgBox Err.Description & Err.Source & " in CTaskDispatcher.NotifyTaskUpdate" & Err.Number, vbCritical
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Private Task Update support Functions
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Private Sub DeliverResults(ByRef strTaskID As String, ByVal lngStatus As Long, ByRef strMemoryKey As String)
On Error GoTo CatchErr
Dim arResultsMemento() As Byte
Dim oRecipient As Object

    If lngStatus <> status_Canceled Then
        arResultsMemento = m_CMemoryMapManager.ReadMapMemory(strMemoryKey)
        m_CBGTRController.ReleaseMemory strMemoryKey
    End If
    
    Set oRecipient = m_colActiveTaskRecipients.Item(strTaskID)
    m_colActiveTaskRecipients.Remove strTaskID
    
    oRecipient.ReturnTaskResults arResultsMemento, lngStatus
    Set oRecipient = Nothing
    
    
    m_lngAvailableTasks = m_lngAvailableTasks + 1
    frmTaskViewer.UpdateTask strTaskID, vbNullString, submitstatus_Terminated
    
Exit Sub
CatchErr:

End Sub

Private Sub ProcessControllerError(ByRef strMemoryKey As String)
On Error GoTo CatchErr
Dim arResultsMemento() As Byte

        arResultsMemento = m_CMemoryMapManager.ReadMapMemory(strMemoryKey)
        m_CMemoryMapManager.ZeroMapMemory strMemoryKey
        RaiseTaskError arResultsMemento
        
Exit Sub
CatchErr:
    MsgBox Err.Description & Err.Source & " raised from CTaskDispatcher.ProcessControllerError" & Err.Number, vbCritical
End Sub


Private Sub Class_Initialize()
On Error GoTo CatchErr
    Set m_colActiveTaskRecipients = New Collection
    Set m_colQueuedTasks = New Collection
    Set m_colQueuedTaskRecipients = New Collection
    Set m_colQueuedTaskIDs = New Collection
    Set m_colMemoryMaps = New Collection
Exit Sub
CatchErr:

End Sub

Private Sub Class_Terminate()
On Error Resume Next
    If m_colActiveTaskRecipients.count > 0 Or m_colQueuedTaskRecipients.count > 0 Then
        MsgBox "The active running tasks must be terminated.   This may require some time." _
        & vbCrLf & "The application will finish automatically, as soon as the worker threads terminate." _
        & vbCrLf & "Click OK to dismiss this dialog.", vbInformation
    End If
    
    m_CBGTRController.Shutdown
    Set m_CBGTRController = Nothing
    Set m_CMemoryMapManager = Nothing
    
End Sub
