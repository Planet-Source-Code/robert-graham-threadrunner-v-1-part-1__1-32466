VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'DLL Object Factory Implementation Code CReportFactory
'- First Release

'Uncomment code for second release version

Option Explicit

Implements IReportFactory

Private Function IReportFactory_EnumerateReports() As Collection
Dim m_Col As Collection
      Set m_Col = New Collection
      With m_Col
           .Add "Summary Report"
           .Add "Detailed Report"
'           .Add "Enhanced Summary Report"
      End With
      Set IReportFactory_EnumerateReports = m_Col
      Set m_Col = Nothing
End Function

Private Function IReportFactory_CreateReport(ByVal lngReportID As Long) As IReport
     Select Case lngReportID

          Case 1
               Set IReportFactory_CreateReport = New CSummary
          Case 2
               Set IReportFactory_CreateReport = New CDetailed
'          Case 3
'               Set IReportFactory_CreateReport = New CEnhancedSummary
     End Select
End Function
