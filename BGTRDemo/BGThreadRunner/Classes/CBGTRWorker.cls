VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBGTRWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************************************
' CBGTRWorker Class Definition
'
' This class defines a worker which understands abstract concepts of
' tasks. The class allows creation of worker threads that make generic
' function calls to a class that defines the specific tasks.
'
' Instancing is set to: 5 - MultiUse
'********************************************************************************************
Option Explicit

Implements ILinkTimer
Private m_CFireTimer As LinkedTimers.CFireTimer
Private m_bIsTaskCancel As Boolean

Private m_CBGTRController As CBGTRController

Private m_strMemoryKey As String
Private m_strTaskID As String

Private m_IBGTRMainFactory As BGTRMainWorks.IBGTRMainFactory
Private m_IBGTRWorkFactory As BGTRMainWorks.IBGTRWorkFactory
Private m_IBGTRWork As BGTRMainWorks.IBGTRWork

Private m_CMemoryMapManager As SharedMemMaps32.CMemoryMapManager
Private m_colMemoryMaps As Collection

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
''^ 'The following functions are public procedures that are
'^  called to initialize a worker and dispatch the worker
'^  on a task
''^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public Sub InitializeWorker(ByVal oController As CBGTRController, ByRef arMemoryMaps() As String)
On Error GoTo CatchErr
Dim i As Long
Dim Low As Long
Dim High As Long
    Set m_CBGTRController = oController
    'Create instance of MemoryMapManager and initialize the memory
    Set m_CMemoryMapManager = New SharedMemMaps32.CMemoryMapManager
    m_CMemoryMapManager.MakeMapFiles arMemoryMaps
    
    'Just in case someone passes us a non-zero base array
    Low = LBound(arMemoryMaps)
    High = UBound(arMemoryMaps)

    'Populate the MemoryMaps collection
    With m_colMemoryMaps
    For i = Low + 1 To High
        .Add arMemoryMaps(i, mapcol_MapKey), arMemoryMaps(i, mapcol_MapKey)
    Next i
    End With
    
Exit Sub
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRWorker.InitializeWorker", Err.Description)
    On Error Resume Next
    Set m_CMemoryMapManager = Nothing
    Set m_CBGTRController = Nothing
    RaiseError strErrorID
End Sub

Public Function DispatchWorker(ByVal strTaskID As String, ByVal strMemoryKey As String) As Long
On Error GoTo CatchErr
    m_strTaskID = strTaskID
    m_strMemoryKey = strMemoryKey
    Call m_CFireTimer.SetNewTimer(Me, "ExecuteTask", ZERO)
    'Reset the Cancel variable
    m_bIsTaskCancel = False
    DispatchWorker = VarPtr(m_bIsTaskCancel)
Exit Function
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRWorker.DispatchWorker", Err.Description)
    On Error Resume Next
    RaiseError strErrorID
End Function

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
''^ 'The following functions are private procedures that are
'^  called by the Worker to execute a task
''^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Private Sub DoExecuteTask()
On Error GoTo CatchErr
Dim pbTemp As PropertyBag
Dim arTaskMemento() As Byte
Dim arWorkMemento() As Byte
Dim arFactoryDescriptors() As String
Dim arWorkDescriptors() As String
Dim Count As Long
Dim i As Long

    'Get the Task Memento
    arTaskMemento = m_CMemoryMapManager.ReadMapMemory(m_strMemoryKey)
    'Unpack the memento
    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = arTaskMemento()
        
        Count = CLng(.ReadProperty("FactoryDescUB"))
        
        ReDim arFactoryDescriptors(0 To Count)
        For i = 0 To Count
            arFactoryDescriptors(i) = .ReadProperty("FactoryDescItem" & i)
        Next i
    
        Count = CLng(.ReadProperty("WorkDescUB"))
        

        ReDim arWorkDescriptors(0 To Count)
        For i = 0 To Count
            arWorkDescriptors(i) = .ReadProperty("WorkDescItem" & i)
        Next i
        arWorkMemento = .ReadProperty("WorkMemento")
        
        
    End With
    Set pbTemp = Nothing
    
    'Get the implementation DLL
    
    Set m_IBGTRWorkFactory = m_IBGTRMainFactory.CreateWorkFactory(arFactoryDescriptors)
    
    'Get the Work class instance that does the actual task
    Set m_IBGTRWork = m_IBGTRWorkFactory.CreateWork(arWorkDescriptors)
    'Deserialize the task
    m_IBGTRWork.SetWorkState arWorkMemento
    'Synchronize the cancel pointers
    SynchCancelPointers
    
    'Ready to roll!
    m_IBGTRWork.InitializeWork
        
    If Not CancelTask Then
        m_IBGTRWork.ExecuteWork
    End If
    
    If Not CancelTask Then
        m_IBGTRWork.TerminateWork
    End If
    
    'All done! or we canceled  Clean up and transfer results
    ConcludeTask
    
Exit Sub
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRWorker.ExecuteTask", Err.Description)
    On Error Resume Next
    Set m_IBGTRWorkFactory = Nothing
    Set m_IBGTRWork = Nothing
    ReturnTaskError strErrorID
End Sub

Private Sub ConcludeTask()
On Error GoTo CatchErr
Dim lngStatus As TaskStatus_enum
Dim arResultsMemento() As Byte
   
    If CancelTask Then
            lngStatus = status_Canceled
    Else
            lngStatus = status_Completed
            arResultsMemento = m_IBGTRWork.WorkResults
            m_CMemoryMapManager.WriteMapMemory m_strMemoryKey, arResultsMemento
    End If
    
    'Destroy these class instances
    Set m_IBGTRWork = Nothing
    Set m_IBGTRWorkFactory = Nothing
    'Notify the Controller and we are all finished!
    m_CBGTRController.NotifyController m_strTaskID, m_strMemoryKey, lngStatus
      
Exit Sub
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRWorker.ConcludeTask", Err.Description)
    On Error Resume Next
    Set m_IBGTRWorkFactory = Nothing
    Set m_IBGTRWork = Nothing
    ReturnTaskError strErrorID
End Sub

Private Sub SynchCancelPointers()
On Error GoTo CatchErr
    m_IBGTRWork.SynchCancelPointers VarPtr(m_bIsTaskCancel)
Exit Sub
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRWorker.SynchCancelPointers", Err.Description)
    On Error Resume Next
    Set m_IBGTRWorkFactory = Nothing
    Set m_IBGTRWork = Nothing
    ReturnTaskError strErrorID
End Sub

Private Function CancelTask() As Boolean
    CancelTask = m_bIsTaskCancel
End Function

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Function to return error info to client via Controller
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub ReturnTaskError(ByRef strErrorID As String)
On Error GoTo CatchErr
Dim pbTemp As PropertyBag

    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "ErrorInfo", ReturnError(strErrorID)
        m_CMemoryMapManager.WriteMapMemory m_strMemoryKey, .Contents
    End With
    Set pbTemp = Nothing
    
   m_CBGTRController.NotifyController m_strTaskID, m_strMemoryKey, status_Aborted
Exit Sub
CatchErr:
'Not much we can do at this point!
On Error Resume Next
    MsgBox "A worker thread could not return error info pertaining to task: " & m_strTaskID _
                & vbCrLf & "This task has failed and system status is unknown!", vbCritical
    'Try to notify controller at least
    m_CBGTRController.NotifyController m_strTaskID, m_strMemoryKey, status_Aborted
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Timer implementations
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub ILinkTimer_FireTimer(ByVal strTimerLink As String, ByVal lngTimerID As Long)
    DoTimerLink strTimerLink, lngTimerID
End Sub

Private Sub DoTimerLink(ByRef strTimerLink As String, ByVal lngTimerID As Long)
On Error GoTo CatchErr

    Select Case strTimerLink

        Case "ExecuteTask"
            DoExecuteTask

        Case Else
            'Raise an Error
            Err.Raise 5 'Invalid Procedure call
    End Select

Exit Sub
CatchErr:
'Should not ever happen, but ..
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRWorker.DoTimerLink", Err.Description)
    ReturnTaskError strErrorID
End Sub

''^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
''^ Class Constructor/Destructer
''^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub Class_Initialize()
    Set m_CFireTimer = New LinkedTimers.CFireTimer
    Set m_IBGTRMainFactory = New CBGTRMainFactory
    Set m_colMemoryMaps = New Collection
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Set m_CBGTRController = Nothing
    Set m_CFireTimer = Nothing

    Set m_IBGTRWork = Nothing
    Set m_IBGTRWorkFactory = Nothing
    Set m_IBGTRMainFactory = Nothing

    Set m_CMemoryMapManager = Nothing
    Set m_colMemoryMaps = Nothing

    MsgBox "Worker thread is terminating"
End Sub
