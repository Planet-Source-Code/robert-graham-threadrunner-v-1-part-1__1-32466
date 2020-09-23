VERSION 5.00
Begin VB.Form frmTestMemento 
   Caption         =   "Testing Memento or Object By Value"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMessage 
      Caption         =   "Description"
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2775
      Begin VB.Label lblMessage 
         Caption         =   "Label1"
         Height          =   2415
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdDestroy 
      Caption         =   "Destroy Objects"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdDepersist 
      Caption         =   "Depersist Objects"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdPersist 
      Caption         =   "Persist Objects"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdPopulate 
      Caption         =   "Populate Objects"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmTestMemento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CStateManager As CStateManager
Private m_strFile As String

Private Sub cmdPopulate_Click()
    PopulateObjects
End Sub

Private Sub cmdPersist_Click()
    PersistObjects
End Sub

Private Sub cmdDestroy_Click()
    DestroyObjects
End Sub

Private Sub cmdDepersist_Click()
    DePersistObjects
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub PopulateObjects()
Dim i As Long
Dim j As Long

Dim oL2A As CL2A
Dim oL3A As CL3A
Dim oL3B As CL3B
Dim oL4A As CL4A
Dim oL4C As CL4C

    Set m_CStateManager = New CStateManager
    With m_CStateManager
        .L1B = New CL1B
        .L1B.Property1 = "Test Level 1B"
        .L1B.Property2 = 1
        
        .L1A = New CL1A
     End With
         
    With m_CStateManager.L1A
        .Property1 = "Test Level 1A"
        .Property2 = 1
        
        For i = 1 To 2
            Set oL3A = New CL3A
                oL3A.L3AProperty1 = 3 * i
                oL3A.L3AProperty2 = "Level 3A Instance_" & i
                oL3A.L4A = New CL4A
                    oL3A.L4A.L4AProperty1 = 4 * i
                    oL3A.L4A.L4AProperty2 = 4 * 2 * i
                    oL3A.L4A.L4AProperty3 = "Level 4A Instance_" & i
                    
            Set oL3B = New CL3B
                oL3B.L4B = New CL4B
                    oL3B.L4B.L4BProperty1 = 4 * i
                    oL3B.L4B.L4BProperty2 = CBool(1 - i)       'Results in one true, one false
                    oL3B.L4Cs = New CL4Cs
                    For j = 1 To 3
                        Set oL4C = oL3B.L4Cs.Add(j)
                        Set oL4C = Nothing
                    Next j
            
            Set oL2A = .L2As.Add(oL3B, oL3A, "Level 2A Instance_" & i, i)
        Next i
    End With
  
End Sub

Private Sub PersistObjects()
Dim baMemento() As Byte
    baMemento = m_CStateManager.GetSuperState
    WriteFile m_strFile, baMemento
    
End Sub

Private Sub DestroyObjects()
    Set m_CStateManager = Nothing
End Sub

Private Sub DePersistObjects()
Dim baMemento() As Byte
    baMemento = ReadFile(m_strFile)
    Set m_CStateManager = New CStateManager
    m_CStateManager.SetSuperState baMemento
End Sub

Private Sub Form_Load()
Dim strMsg As String

    'Create a file to save state data:
    m_strFile = App.Path & "\" & "memento.dat"
    

    strMsg = "This project demonstrates the Object By Value technique." & vbCrLf & vbCrLf _
    & "  The button clicks will: " & vbCrLf _
    & "    - create a small set of objects" & vbCrLf _
    & "    - persist their state to a file" & vbCrLf _
    & "    - destroy the original set" & vbCrLf _
    & "    - recreate them from the file" & vbCrLf & vbCrLf _
    & "To see the effects, run the program in debug mode, " _
    & " and step through the sequence. "
    
    Me.lblMessage.Caption = strMsg
End Sub
