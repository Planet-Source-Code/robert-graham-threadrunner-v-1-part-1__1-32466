VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ILinkTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************************************
' ILinkTimer Interface Definition
'
' This class defines an interface which is implemented by the client
' requesting a timer. The single function returns the string passed
' by the client which is used by the client to interpret how to process the
' timer fire event.
'
' Instancing is set to: 2 - PublicNotCreatable
'********************************************************************************************
Public Sub FireTimer(ByVal strTimerLink As String, ByVal lngTimerID As Long)
End Sub


