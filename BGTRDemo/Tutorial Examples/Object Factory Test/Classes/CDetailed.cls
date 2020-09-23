VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDetailed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module CDetailed
Option Explicit

'Implements our standard client interface
Implements IReport

Private Sub IReport_DoSomeReport()
    DoDetailedReport
End Sub

Private Sub DoDetailedReport()
     'Specific implementation code goes here.
     MsgBox "Print a Detailed Report!"
End Sub

