VERSION 5.00
Begin VB.MDIForm frmMDIMain 
   BackColor       =   &H80000001&
   Caption         =   "BGThreadRunner Demonstration"
   ClientHeight    =   7815
   ClientLeft      =   315
   ClientTop       =   630
   ClientWidth     =   11385
   LinkTopic       =   "MDIForm1"
   Begin VB.Timer tmrInstall 
      Interval        =   1000
      Left            =   0
      Top             =   840
   End
   Begin VB.Timer tmrMain 
      Left            =   480
      Top             =   840
   End
   Begin VB.PictureBox pctMDIMain 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   11385
      TabIndex        =   0
      Top             =   0
      Width           =   11385
      Begin VB.TextBox txtThread 
         Height          =   285
         Left            =   7440
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblMainFore 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "BGThreadRunner Demonstration"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   540
         Left            =   480
         TabIndex        =   1
         Top             =   20
         Width           =   8055
      End
      Begin VB.Label lblMainBack 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "BGThreadRunner Demonstration"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   540
         Left            =   520
         TabIndex        =   2
         Top             =   60
         Width           =   8055
      End
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_DemoSamples 
      Caption         =   "&Demo Samples"
      Begin VB.Menu mnu_RecordsetProcess 
         Caption         =   "&Recordset Process"
      End
      Begin VB.Menu mnu_QueryRunner 
         Caption         =   "&Query Runner"
      End
   End
   Begin VB.Menu mnu_Tasks 
      Caption         =   "&Task Status"
      Begin VB.Menu mnu_TaskViewer 
         Caption         =   "Task &Viewer"
      End
   End
   Begin VB.Menu mnu_Windows 
      Caption         =   "&Windows"
      Index           =   0
      WindowList      =   -1  'True
      Begin VB.Menu mnu_WindowsItems 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_DemoHelp 
         Caption         =   "&Demo Help"
      End
      Begin VB.Menu mnu_ThreadRunnerHelp 
         Caption         =   "&ThreadRunner Help"
      End
      Begin VB.Menu mnu_HelpSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_AboutThreadRunner 
         Caption         =   "&About ThreadRunner"
      End
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************************
'   frmMDIMain Form Definition
'
'   Main MDI form for the app, with a hidden textbox in the picturebox window.
'
'********************************************************************************************
Option Explicit

'Constant required for the About dialog
Private Const IDM_ABOUT As Long = 1010
              
'Enum definition for Windows menu commands
Private Enum WindowsActions_enum
    winmnu_WindowsTileVertically = 0
    winmnu_WindowsTileHorizontally = 1
    winmnu_WindowsCascade = 2
    winmnu_WindowsArrange = 3
    winmnu_WindowsClose = 4
    winmnu_WindowsCloseAll = 5
    winmnu_WindowsItems = 6
End Enum

'Collection to temporarily cache Controller messages
Private colMessages As Collection

'Hwnd for help file - not used actively
Private m_hwndHelp As Long

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ The following functions capture and reconcile all Controller
'^ notifications
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'This is the key function which causes the app to respond to Controller
'notifications sent via SendMessage
Private Sub txtThread_Change()
Dim strMsg As String

    If Len(Me.txtThread) > 0 Then
        strMsg = Me.txtThread.Text
        'Cache the message and set a timer,
        'so the call will immediately return to the Controller
        colMessages.Add strMsg
        tmrMain.Interval = 10
        'Clear the hidden window
        Me.txtThread = vbNullString
    End If
    
End Sub

'This is an ordinary VB timer, used to asynchronously
'process Controller notifications
Private Sub tmrMain_Timer()
    'Kill the timer
    tmrMain.Interval = 0
    'Process the Controller notification
    ProcessMessages
End Sub

'This function processes the Controller notifications
Private Sub ProcessMessages()
Dim strMsg As String

    With colMessages
        If .count > 0 Then
            strMsg = .Item(FIRST_ITEM)
            .Remove (FIRST_ITEM)
            g_CTaskDispatcher.NotifyTaskUpdate strMsg
            If .count > 0 Then
                'Set another timer to repeat this loop
                tmrMain.Interval = 10
            End If
        End If
    End With

End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Main Menu functions
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'File Menu commands
Private Sub mnu_Exit_Click()
  Unload Me
End Sub

'Demo Samples Menu commands
Private Sub mnu_QueryRunner_Click()
    'This is a single instance form - display the active form
    'if user clicks on the menu item multiple times
    frmQueryRunner.Show
    SetFormPosition frmQueryRunner, vbTop
End Sub
Private Sub mnu_RecordsetProcess_Click()
    'This is a single instance form - display the active form
    'if user clicks on the menu item multiple times
    frmMakeRecords.Show
    SetFormPosition frmMakeRecords, vbTop
End Sub

'Task StatusMenu Commands
Private Sub mnu_TaskViewer_Click()
    frmTaskViewer.Show
    SetFormPosition frmTaskViewer, vbTop
End Sub

'Help Menu Commands
Private Sub mnu_DemoHelp_Click()
Dim strDemoPath As String
Dim strTest As String
Dim strHelpPath As String

  strDemoPath = App.Path
  Do While strTest <> "\"
      strTest = Right(strDemoPath, 1)
      strDemoPath = Left(strDemoPath, Len(strDemoPath) - 1)
  Loop
  strHelpPath = strDemoPath & "\" & "BGTR Help"
       
  App.HelpFile = strHelpPath & "\BGTRDemo.chm"
  m_hwndHelp = HtmlHelp(Me.hwnd, App.HelpFile, HH_DISPLAY_TOPIC, ByVal "about_the_bgthreadrunner_demo_ap.htm")
End Sub

Private Sub mnu_ThreadRunnerHelp_Click()
Dim strDemoPath As String
Dim strTest As String
Dim strHelpPath As String

  strDemoPath = App.Path
  Do While strTest <> "\"
      strTest = Right(strDemoPath, 1)
      strDemoPath = Left(strDemoPath, Len(strDemoPath) - 1)
  Loop
  App.HelpFile = strDemoPath & "\BGThreadRunnerDemoHelp.chm"
  m_hwndHelp = HtmlHelp(Me.hwnd, App.HelpFile, HH_DISPLAY_TOPIC, ByVal "cover.htm")
End Sub

'About Menu Commands
Private Sub mnu_AboutThreadRunner_Click()
    frmAbout.Show vbModal
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Main Menu Support functions
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Windows Menu Commands
Private Sub mnu_WindowsItems_Click(Index As Integer)
Dim frm As Form

    Select Case Index
    
    Case winmnu_WindowsTileVertically
         Me.Arrange vbTileVertical
      
      Case winmnu_WindowsTileHorizontally
         Me.Arrange vbTileHorizontal
         
      Case winmnu_WindowsCascade
         Me.Arrange vbCascade
         
      Case winmnu_WindowsArrange
         Me.Arrange vbArrangeIcons
         
      ' Close active form
      Case winmnu_WindowsClose
         CloseMDIForm
         
      ' Close all windows
      Case winmnu_WindowsCloseAll
         CloseAllMDIForms
    
   End Select
Exit Sub
CatchErr:

End Sub

Private Sub CreateWindowsMenu()
Dim i As Long
    
    'Execute a load to dynamically load all the items in the Windows menu
    For i = winmnu_WindowsTileVertically + 1 To winmnu_WindowsCloseAll
        Load mnu_WindowsItems(i)
    Next i
    'Now set the caption for each menu item dynamically loaded
    mnu_WindowsItems(winmnu_WindowsTileVertically).Caption = "Tile &Vertically"
    mnu_WindowsItems(winmnu_WindowsTileHorizontally).Caption = "Tile &Horizontally"
    mnu_WindowsItems(winmnu_WindowsCascade).Caption = "&Cascade"
    mnu_WindowsItems(winmnu_WindowsArrange).Caption = "Arrange &Icons"
    mnu_WindowsItems(winmnu_WindowsClose).Caption = "&Close"
    mnu_WindowsItems(winmnu_WindowsCloseAll).Caption = "Close &All"
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Form Class Constructor/Destructor
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Private Sub MDIForm_Load()
On Error GoTo CatchErr
    'Call function to dynamically create the Windoows menu list
    CreateWindowsMenu
    'Instantiate a message collection, and a global Dispatcher
    Set colMessages = New Collection
    Set g_CTaskDispatcher = New CTaskDispatcher
    'Initialize the Dispatcher - creates memory maps,etc.
    g_CTaskDispatcher.InitializeTaskService
    
Exit Sub
CatchErr:
Dim strErrorID As String
    strErrorID = SaveError(Err.Number, Err.Source & " in frmMDIMain_Load", Err.Description)
    Set g_CTaskDispatcher = Nothing
    RaiseError strErrorID
    MsgBox "the application will now terminate due to fatal errors!", vbCritical
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Set g_CTaskDispatcher = Nothing
End Sub

