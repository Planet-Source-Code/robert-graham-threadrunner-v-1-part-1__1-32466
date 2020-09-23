VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBGTRWorkFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************************************
' IBGTRWorkFactory Interface Definition
'
' This class defines an interface which allows a Worker thread
' to request concrete class instances of the classes to execute
' a background task.
'
' Instancing is set to: 2 - PublicNotCreatable
'********************************************************************************************
Option Explicit

Public Function CreateWork(ByRef arWorkDescriptors() As String) As IBGTRWork
End Function


