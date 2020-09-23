VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBGTRMainFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************************************
' IBGTRMainFactory Interface Definition
'
' This class defines an interface which allows a Worker thread
' to request an implementation instance of the secondary class
' factory which will in turn provide the actual task impolmentation classes
'
' Instancing is set to: 2 - PublicNotCreatable
'********************************************************************************************
Option Explicit

Public Function CreateWorkFactory(ByRef arFactoryDescriptors() As String) As IBGTRWorkFactory
End Function

