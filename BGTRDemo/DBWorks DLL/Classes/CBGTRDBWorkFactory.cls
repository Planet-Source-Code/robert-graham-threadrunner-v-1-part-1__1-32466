VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBGTRDBWorkFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CBGTRDBWorkFactory Class Definition
'   Class defines secondary concrete factory class for the DBWorks DLL
'   This factory provides actual tasks by creating the appropriate concrete
'   class instances implementing the IBGTRWork interface
'
'   Instancing is set to:  5 - MultiUse
''
'********************************************************************************************
Option Explicit

Implements IBGTRWorkFactory

Private Function IBGTRWorkFactory_CreateWork(arWorkDescriptors() As String) As IBGTRWork
On Error GoTo CatchErr
    
    Select Case arWorkDescriptors(0)
        Case "Query"
            Set IBGTRWorkFactory_CreateWork = New CBGTRUserQuery
            
        Case "Recordset"
            Set IBGTRWorkFactory_CreateWork = New CBGTRMakeRecords
    End Select

Exit Function
CatchErr:
    Err.Raise Err.Number, Err.Source & " in CBGTRDBWorks.CBGTRWorkFactory.IBGTRWorkFactory_CreateWork", Err.Description
End Function
