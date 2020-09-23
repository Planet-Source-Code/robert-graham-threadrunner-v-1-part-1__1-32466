VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IReportFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Class Module IReportFactory
'DLL Object Factory Interface IReportFactory– does not ever change
Option Explicit

Public Function CreateReport(ByVal lngReportID As Long) As IReport
End Function

Public Function EnumerateReports() As Collection
End Function

