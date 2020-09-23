VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLinkTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************************
' CLinkTimer Class Definition
'
' This class defines an object which is created by for each Win32 timer
' object. This class saves the client information passed in with a timer
' request and defines the functions needed to call back into the client
' when the timer fires.
'
' Instancing is set to: 1 - Private
'********************************************************************************************
Option Explicit

'Declare module level variables to save the link info.
Private m_oTimer As ILinkTimer
Private m_lngTimerID As Long
Private m_strTimerLink As String

Public Property Let TimerID(ByVal lngTimerID As Long)
    m_lngTimerID = lngTimerID
End Property
Public Property Get TimerID() As Long
    TimerID = m_lngTimerID
End Property

Public Property Let TimerLink(ByVal strLinkTimer As String)
    m_strTimerLink = strLinkTimer
End Property
Public Property Get TimerLink() As String
    TimerLink = m_strTimerLink
End Property

Public Property Let Timer(ByRef p_objTimer As ILinkTimer)
    Set m_oTimer = p_objTimer
End Property
'This function actually calls the client back
Public Sub FireTimer()
    m_oTimer.FireTimer m_strTimerLink, m_lngTimerID
    Set m_oTimer = Nothing
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Set m_oTimer = Nothing
End Sub


