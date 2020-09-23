VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBGTRWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************************************
' IBGTRWork Interface Definition
'
' This class defines an interface which allows a Worker thread
' to call the functions in a concrete class instance in order to execute
' a background task.
'
' Instancing is set to: 2 - PublicNotCreatable
'********************************************************************************************
Option Explicit

Public Property Get WorkResults() As Byte()
End Property

Public Sub SynchCancelPointers(ByVal pCancel As Long)
End Sub

Public Sub InitializeWork()
End Sub

Public Sub ExecuteWork()
End Sub

Public Sub TerminateWork()
End Sub

Public Sub SetWorkState(ByRef arWorkMemento() As Byte)
End Sub

Public Function GetWorkState() As Byte()
End Function
