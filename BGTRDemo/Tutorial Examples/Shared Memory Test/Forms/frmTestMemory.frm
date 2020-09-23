VERSION 5.00
Begin VB.Form frmTestMemory 
   Caption         =   "Shared Memory Test Application"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDropSemaphore 
      Caption         =   "Drop Semaphore"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdHoldSemaphore 
      Caption         =   "Hold Semaphore"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdDropMutex 
      Caption         =   "Drop Mutex"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdHoldMutex 
      Caption         =   "Hold Mutex"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read Memory"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write Memory"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtTest 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton cmdMap 
      Caption         =   "Map Memory"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblTest 
      Caption         =   "Enter a test string:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmTestMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CMemoryMapFileTest As CMemoryMapFileTest
'Just declare a const string for the name
Private Const MAP_NAME As String = "SharedMem32Test"
'Just declare a const long for the size - 50KB
Private Const MAP_SIZE As Long = 50



Private Sub cmdMap_Click()
On Error GoTo CatchErr
   
   m_CMemoryMapFileTest.MapMemory MAP_NAME, MAP_SIZE
   'Don't map it twice!!!!!!!!!!!!!!
   cmdMap.Enabled = False       'Turn it off and leave it off!

Exit Sub
CatchErr:
    MsgBox Err.Number & "  " & Err.Source & vbCrLf & Err.Description, vbCritical
End Sub





Private Sub cmdWrite_Click()
On Error GoTo CatchErr
Dim pbTemp As PropertyBag
Dim arData() As Byte
Dim strTest As String

    strTest = txtTest
    
    Set pbTemp = New PropertyBag
    With pbTemp
        .WriteProperty "TestString", strTest
        arData = .Contents
    End With
    Set pbTemp = Nothing
    
    m_CMemoryMapFileTest.WriteMemory arData
    
Exit Sub
CatchErr:
    MsgBox Err.Number & "  " & Err.Source & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub cmdRead_Click()
On Error GoTo CatchErr
Dim pbTemp As PropertyBag
Dim arData() As Byte
    
    arData = m_CMemoryMapFileTest.ReadMemory
    
    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = arData
        txtTest = .ReadProperty("TestString")
        Beep
    End With
    Set pbTemp = Nothing

Exit Sub
CatchErr:
    MsgBox Err.Number & "  " & Err.Source & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub cmdHoldMutex_Click()
On Error GoTo CatchErr
    m_CMemoryMapFileTest.HoldMutex
Exit Sub
CatchErr:
    MsgBox Err.Number & "  " & Err.Source & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub cmdDropMutex_Click()
On Error GoTo CatchErr
    m_CMemoryMapFileTest.DropMutex
    
Exit Sub
CatchErr:
    MsgBox Err.Number & "  " & Err.Source & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub cmdHoldSemaphore_Click()
On Error GoTo CatchErr
    m_CMemoryMapFileTest.HoldSemaphore
Exit Sub
CatchErr:
    MsgBox Err.Number & "  " & Err.Source & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub cmdDropSemaphore_Click()
On Error GoTo CatchErr
    
    m_CMemoryMapFileTest.DropSemaphore
    
Exit Sub
CatchErr:
    MsgBox Err.Number & "  " & Err.Source & vbCrLf & Err.Description, vbCritical
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set m_CMemoryMapFileTest = New CMemoryMapFileTest
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_CMemoryMapFileTest = Nothing
End Sub
