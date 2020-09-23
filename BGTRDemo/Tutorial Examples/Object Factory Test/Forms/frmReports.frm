VERSION 5.00
Begin VB.Form frmReports 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReport 
      Caption         =   "Print Report"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.ListBox lstReports 
      Height          =   1815
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Calling code in client that is completely generic – does not ever change

'Declare object variables
Private m_IReportFactory As ReportsDLL.IReportFactory
Private m_IReport As ReportsDLL.IReport

Private Sub Form_Load()
    'Create an instance of the concrete ReportFactory
    Set m_IReportFactory = New ReportsDLL.CReportFactory
    'Call function to populate the listbox
    PopulateReportList
End Sub

'Populate listbox from the DLL.
Private Sub PopulateReportList()
Dim col As Collection
    Set col = m_IReportFactory.EnumerateReports
          With frmReports.lstReports
                For Each Item In col
                      .AddItem Item
               Next Item
          End With
    Set col = Nothing
End Sub

Private Sub cmdReport_Click()
    'Call the factory to create a concrete Report instance based on the selection
    Set m_IReport = m_IReportFactory.CreateReport(lstReports.ListIndex + 1)
    'Call the DoSomeReport method
    m_IReport.DoSomeReport
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_IReport = Nothing
    Set m_IReportFactory = Nothing
End Sub
