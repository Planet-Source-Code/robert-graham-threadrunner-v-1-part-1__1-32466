VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFireTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************************************
' CFireTimer Class Definition
'
' This class defines an object which is created by a client
' and defines the functions needed by the client to create and use Win32
' fire once timers. The class instantiates a collection of CLinkTimer objects
' but the collection itself is global to the DLL.
'
' Instancing is set to: 5 - MultiUse
'********************************************************************************************
Option Explicit

'Define a public enum for the no timer error so it may be published
'to any clients

Public Enum LinkedTimersErrorCodes_enum
    linkedtimerserr_NoTimersAvailable = 40000
End Enum

Public Function SetNewTimer(ByRef p_objTimer As ILinkTimer, strLinkTimer As String, Optional ByVal lngMilliSeconds As Long = 0) As Long
On Error GoTo CatchErr
Dim objCLinkTimer As CLinkTimer
Dim lngTimerID As Long
    lngTimerID = StartTimer(lngMilliSeconds)
            If lngTimerID <> 0 Then
                Set objCLinkTimer = New CLinkTimer
                With objCLinkTimer
                    .Timer = p_objTimer
                    .TimerID = lngTimerID
                    .TimerLink = strLinkTimer
                End With
                g_PrimedTimers.Insert objCLinkTimer, Str$(lngTimerID)
                Set objCLinkTimer = Nothing
                SetNewTimer = lngTimerID
                lngTimerID = 0
            End If
Exit Function
CatchErr:
    Dim strErrorID As String
    strErrorID = SaveError(Err.Number, "LinkedTimers.CFireTimer.SetNewTimer", Err.Description)
    On Error Resume Next
    'Destroy any timer objects which may have been created
    ClearTimer lngTimerID
    Set objCLinkTimer = Nothing
    'Raise an error in client
    RaiseError strErrorID
End Function

Public Sub KillTimer(ByVal lngTimerID As Long)
On Error Resume Next
    ClearTimer lngTimerID
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Class Constructor/Destructor
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Private Sub Class_Initialize()
    'Create global g_PrimedTimers collection
    If g_PrimedTimers Is Nothing Then
        Set g_PrimedTimers = New CLinkTimers
    End If
End Sub

Private Sub Class_Terminate()
On Error Resume Next
Dim oCLinkTimer As CLinkTimer
    'Make sure all active timers are destroyed
    For Each oCLinkTimer In g_PrimedTimers
        ClearTimer oCLinkTimer.TimerID
        Set oCLinkTimer = Nothing
    Next oCLinkTimer
    Set g_PrimedTimers = Nothing
End Sub


