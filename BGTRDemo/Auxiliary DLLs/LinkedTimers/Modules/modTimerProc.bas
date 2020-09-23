Attribute VB_Name = "modTimerProc"
'********************************************************************************************
' modTimerProc Module Definition
'
' This is a standard module that calls the API declarations to create
' Win32 fire once timers and provides the callback entry function for
' the timer created.
'
' All functions which act in the Win32 domain with the physical timer
' reside in this module.
'********************************************************************************************
Option Explicit

'Const for timer hwnd - these timers will always use zero
Public Const HWND_NONE = 0
'Const for timer IDEvent - these timers are callbacks and will always use zero
Public Const NIDEVENT_NONE = 0

Public Function StartTimer(lngMilliSeconds As Long) As Long
    StartTimer = SetTimer(HWND_NONE, NIDEVENT_NONE, lngMilliSeconds, AddressOf TimerProc)
     'Raise error if no timer was created.
     If StartTimer = 0 Then _
       Err.Raise vbObjectError + linkedtimerserr_NoTimersAvailable, _
                       " in modMainLinkedTimers.StartTimer", LoadResString(linkedtimerserr_NoTimersAvailable)
End Function

Public Sub TimerProc(ByVal hwndOwner As Long, _
                                 ByVal lngMsg As Long, _
                                 ByVal lngTimerID As Long, _
                                 ByVal lngTime As Long)
 'Set Error Resume Next because TimerProc may be called after objects are destroyed.
On Error Resume Next
Dim oTimerLink As CLinkTimer

    'Retrieve the TimerLink object
    Set oTimerLink = g_PrimedTimers.Item(Str$(lngTimerID))
    'Destroy the timer
    ClearTimer lngTimerID
    'Fire the primer link and clean up
    oTimerLink.FireTimer
    Set oTimerLink = Nothing
End Sub

Public Sub ClearTimer(ByVal lngTimerID As Long)
 'Set Error Resume Next because ClearTimer may be called after objects are destroyed.
On Error Resume Next
    KillTimer HWND_NONE, lngTimerID
    g_PrimedTimers.Remove Str$(lngTimerID)
End Sub


