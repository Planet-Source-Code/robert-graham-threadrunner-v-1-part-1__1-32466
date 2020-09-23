VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBGTRWorkerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************************
' CBGTRWorkerInfo Class Definition
'
' This class defines a WorkerInfo object which creates Worker threads and
' provides most of the functionality used by the Controller to dispatch them
' to perform background tasks.
'
' Instancing is set to: 1 - Private
'********************************************************************************************
Option Explicit

Private m_CBGTRWorker As CBGTRWorker
Private m_pCancel As Long
Private m_strMemoryKey As String

Public Property Get Worker() As CBGTRWorker
    Set Worker = m_CBGTRWorker
End Property

Public Property Let MemoryKey(ByVal strMemoryKey As String)
    m_strMemoryKey = strMemoryKey
End Property
Public Property Get MemoryKey() As String
    MemoryKey = m_strMemoryKey
End Property

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ The following functions create and initialize Worker
'^ objects
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public Function CreateWorker(ByVal oController As CBGTRController, ByRef arMemoryMaps() As String) As Long
On Error GoTo CatchErr

    'Create a new worker thread
    Set m_CBGTRWorker = CreateObject("BGThreadRunner.CBGTRWorker")
    'Set the Controller reference
    m_CBGTRWorker.InitializeWorker oController, arMemoryMaps
    'Return the ObjPtr to the Controller, used for a unique ID only
    'Note that this value is not passed over an apartment boundary, so this is a valid ObjPtr
    CreateWorker = ObjPtr(m_CBGTRWorker)
Exit Function
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRWorkerInfo.CreateWorker", Err.Description)
    On Error Resume Next
    Set m_CBGTRWorker = Nothing
    RaiseError strErrorID
End Function

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ The following public functions are called by the Controller
'^ to dispatch Workers on a task or cancel a task
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Public Sub DispatchWorker(ByVal strTaskID As String, ByVal strMemoryKey As String)
On Error GoTo CatchErr
    'Save this in the local variable
    m_strMemoryKey = strMemoryKey
    'Make sure our cancel pointer is valid
    m_pCancel = m_CBGTRWorker.DispatchWorker(strTaskID, strMemoryKey)
Exit Sub
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRWorkerInfo.DispatchWorker", Err.Description)
    On Error Resume Next
    RaiseError strErrorID
End Sub

'This needs to be Public now!
Public Sub DoCancelTask()
On Error GoTo CatchErr
    InterlockedIncrement m_pCancel
Exit Sub
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in CBGTRWorkerInfo.DoCancelTask", Err.Description)
    On Error Resume Next
    RaiseError strErrorID
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Class constructor/destructor
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub Class_Initialize()
    'Initialize our variables just for good measure
    m_pCancel = 0
    m_strMemoryKey = vbNullString
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    DoCancelTask
    Set m_CBGTRWorker = Nothing
    MsgBox "WorkerInfo is terminating!"
End Sub


