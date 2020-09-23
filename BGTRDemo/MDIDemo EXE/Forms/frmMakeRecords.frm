VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMakeRecords 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recordset Processing Sample Function"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3660
   ScaleWidth      =   5745
   Begin VB.Timer tmrBlink 
      Left            =   120
      Top             =   3000
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Recordset Progress"
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   5295
      Begin VB.TextBox txtTotal 
         Height          =   315
         Left            =   3045
         TabIndex        =   0
         ToolTipText     =   "Read/Write:  Total number of records to be created"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtCurrent 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "The value displayed is updated periodically to indicate task progress"
         Top             =   840
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar pbarMakeRecords 
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblRecords 
         Alignment       =   1  'Right Justify
         Caption         =   "records"
         Height          =   255
         Left            =   4320
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblOf 
         Alignment       =   2  'Center
         Caption         =   "of"
         Height          =   255
         Left            =   2555
         TabIndex        =   6
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblCurrent 
         Caption         =   "Created"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Execute"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      ToolTipText     =   "Click to execute task"
      Top             =   3120
      Width           =   1215
   End
   Begin MSComCtl2.Animation anmShowAction 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1931
      _Version        =   393216
      Center          =   -1  'True
      FullWidth       =   369
      FullHeight      =   73
   End
End
Attribute VB_Name = "frmMakeRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private constants for location of hidden window

Private Const RECORDS_FRAME_CLASS As String = "ThunderRT6Frame"
Private Const RECORDS_TEXTBOX_CLASS As String = "ThunderRT6TextBox"

'Stub out the above and use these instead in order to run in the IDE
'Private Const RECORDS_FRAME_CLASS As String = "ThunderFrame"
'Private Const RECORDS_TEXTBOX_CLASS As String = "ThunderTextBox"


Private m_hwndUpdateProgress As Long

Private Const DefaultHeight As Long = 4050
Private Const DefaultWidth As Long = 5865

Private strAVIFile As String
Private bIsSwitch As Boolean

'Cache the caret blink rate
Private lngBlinkRate As Long
'Cache last state
Private lngLastBlinkState As Long
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ 'This is the function required by the Dispatcher to
'^  return results from task execution
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public Sub ReturnTaskResults(ByRef arResultsMemento() As Byte, ByVal lngStatus As Long)
On Error GoTo CatchErr

    Select Case lngStatus
        
        Case status_Canceled
            txtCurrent = "Canceled"
        
        Case status_Completed
            txtCurrent = "Completed"
       
        Case status_Aborted
            RaiseTaskError arResultsMemento
            
    End Select
    
        cmdRun.Enabled = True
        bIsSwitch = False
        pbarMakeRecords.Value = 0
        On Error Resume Next
        anmShowAction.Stop
        anmShowAction.Visible = False
        CheckActiveStatus
 Exit Sub
CatchErr:
    MsgBox "Error Number: " & Err.Number & " Source: " & Err.Source & " raised from frmQueryViewer.ReturnTaskResults" & vbCrLf & Err.Description, vbCritical
End Sub
Private Sub CheckActiveStatus()
Dim frm As Form

    Set frm = frmMDIMain.ActiveForm
    
    If Not frm Is Nothing Then
        If frm.Caption = Me.Caption Then
            'Do Nothing
        Else
            Me.tmrBlink.Interval = lngBlinkRate
        End If
        Set frm = Nothing
    Else
        Me.tmrBlink.Interval = lngBlinkRate
    End If
         
End Sub

Private Sub Form_Activate()
 
 If tmrBlink.Interval > 0 Then
    tmrBlink.Interval = 0
    lngLastBlinkState = FlashWindow(Me.hwnd, INVERT)
 
    If lngLastBlinkState = 0 Then
        lngLastBlinkState = FlashWindow(Me.hwnd, INVERT)
    End If
End If

End Sub


Private Sub RunMakeRecords()
On Error GoTo CatchErr
Dim strTaskID As String

    strTaskID = g_CTaskDispatcher.SubmitTask(MakeRecordTask, frmMakeRecords, "Execute Recordset Processing")
    
Exit Sub
CatchErr:
    Err.Raise Err.Number, Err.Source & " in frmMakeRecords.RunMakeRecords", Err.Description
End Sub
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Configure Functions - Create class instances
'^ to configure for task execution
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Function ConfigureMakeRecords() As Byte()
On Error GoTo CatchErr
Dim oMakeRecord As CBGTRMakeRecords
    
    Set oMakeRecord = New CBGTRMakeRecords
    
    With oMakeRecord
        'Just hardcore these - it is a demo only after all!
        .Connection.Server = App.Path & "\db\BGTRDemoJetDB.mdb"
        .Connection.User = "Admin"
        .Connection.Password = ""
        .FinalRecordCount = CLng(txtTotal)
        .UpdateHwnd = m_hwndUpdateProgress
        ConfigureMakeRecords = .GetSuperState
    End With
    
    Set oMakeRecord = Nothing
    
Exit Function
CatchErr:
     Err.Raise Err.Number, Err.Source & " in frmMakeRecords.ConfigureMakeRecords", Err.Description
End Function

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Task functions create a task memento based upon
'^ the number of records to create for submission as a
'^ background task
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Private Function MakeRecordTask() As Byte()
On Error GoTo CatchErr
Dim arFactoryDescriptor() As String
Dim arWorkDescriptor() As String
Dim pbTemp As PropertyBag
Dim i As Long
    
    ReDim arFactoryDescriptor(0)
    arFactoryDescriptor(0) = "DBTasks"

    ReDim arWorkDescriptor(0)
    arWorkDescriptor(0) = "Recordset"

    Set pbTemp = New PropertyBag
    
    With pbTemp
        .WriteProperty "FactoryDescUB", UBound(arFactoryDescriptor)
        For i = 0 To UBound(arFactoryDescriptor)
        .WriteProperty "FactoryDescItem" & i, arFactoryDescriptor(i)
        Next i
        
        .WriteProperty "WorkDescUB", UBound(arWorkDescriptor)
        For i = 0 To UBound(arWorkDescriptor)
        .WriteProperty "WorkDescItem" & i, arWorkDescriptor(i)
        Next i
        
        .WriteProperty "WorkMemento", ConfigureMakeRecords
        
        MakeRecordTask = .Contents
    End With

Exit Function
CatchErr:
     Err.Raise Err.Number, Err.Source & " in frmMakeRecords.MakeRecordTask", Err.Description
End Function

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Major form action command buttons
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub cmdRun_Click()
    lblCurrent.Caption = "Created"
    txtCurrent = vbNullString
    RunMakeRecords
    cmdRun.Enabled = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub tmrBlink_Timer()
    lngLastBlinkState = FlashWindow(Me.hwnd, INVERT)
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Form controls
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub txtCurrent_Change()
On Error GoTo CatchErr
    If txtCurrent = "Update" Then
        If Not bIsSwitch Then
            pbarMakeRecords.Value = 0
            txtCurrent = vbNullString
            bIsSwitch = True
        End If
        
        lblCurrent.Caption = "Updated"
        
        strAVIFile = App.Path & "\filemove.avi"
        anmShowAction.Open strAVIFile
        anmShowAction.Visible = True
        anmShowAction.Play
    End If
    pbarMakeRecords.Value = txtCurrent / txtTotal * 100
Exit Sub
CatchErr:

End Sub


'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Various support functions
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'This function determines hwnd for progress report window
Private Sub FindProgressWindow()
On Error GoTo CatchErr
Dim hwnd_Frame As Long
Dim hwnd_Text As Long

    hwnd_Frame = FindWindowEx(Me.hwnd, 0, RECORDS_FRAME_CLASS, Me.fraStatus.Caption)
    
    'Facilitates finding the correct one!
    Me.txtCurrent = "Find Me"
    
    m_hwndUpdateProgress = FindWindowEx(hwnd_Frame, 0, RECORDS_TEXTBOX_CLASS, Me.txtCurrent.Text)
    Me.txtCurrent = vbNullString
    
Exit Sub
CatchErr:
    Err.Raise Err.Number, Err.Source & " in frmMakeRecords.FindProgressWindow", Err.Description
End Sub

'Cannot load AVI files directly from Res file without a hassle
'and cannot copy with the more efficient methods, so must resort
'to 1-byte-at-a-time copy method
Private Sub LoadAVIFile()
On Error GoTo CatchErr
Dim baAVIFile() As Byte
Dim strAVIFile As String
Dim i As Long
Dim lngAVISize As Long

    strAVIFile = App.Path & "\filemove.avi"
    If Not FileExists(strAVIFile) Then
       Open strAVIFile For Binary As #1
        baAVIFile = LoadResData(RES_FILEMOVE_AVI, "CUSTOM")
        lngAVISize = UBound(baAVIFile)
        For i = 0 To lngAVISize
            Put #1, , baAVIFile(i)
        Next i
    
        Close #1
    End If
Exit Sub
CatchErr:

End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Form Class Constructor/Destructor
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub Form_Load()
    Me.Width = DefaultWidth
    Me.Height = DefaultHeight
    CenterForm Me
    g_FormsCount = g_FormsCount + 1
    'Set default record count
    Me.txtTotal = 5000
    FindProgressWindow
    LoadAVIFile
    lngBlinkRate = GetCaretBlinkTime
End Sub

Private Sub Form_Unload(Cancel As Integer)
    g_FormsCount = g_FormsCount - 1
End Sub


