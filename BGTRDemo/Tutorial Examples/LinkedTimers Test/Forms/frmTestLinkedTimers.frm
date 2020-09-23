VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTestLinkedTimers 
   Caption         =   "Test LinkedTimers DLL"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSetTimers 
      Caption         =   "Set Timers"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtTime 
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox cboLink 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   1200
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid msfgTimers 
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdKillTimer 
      Caption         =   "Kill Timer"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddTimer 
      Caption         =   "Add Timer"
      Height          =   375
      Left            =   293
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblTime 
      Caption         =   "Timer Interval (ms):"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblLink 
      Caption         =   "Timer Link String:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "frmTestLinkedTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ILinkTimer

Private m_CFireTimer As LinkedTimers.CFireTimer

Private m_lngTimerID As Long

Sub AddTimerToGrid(ByVal lngTimerID As Long)
   With msfgTimers
   .Row = 1
   .Col = 0
    If Len(.Text) = 0 Then
        .Text = cboLink
    Else
        .AddItem cboLink
    End If
    .Row = .Rows - 1
    .Col = 1
     .Text = txtTime
    .Col = 2
    .Text = CStr(lngTimerID)
End With

End Sub
Sub RemoveTimerFromGrid(ByVal lngTimerID As Long)
Dim i As Long
    With msfgTimers
        If .Rows = 2 Then
            .Row = 1
            For i = 0 To 2
                .Col = i
                .Text = vbNullChar
            Next i
        Else
            .Col = 2
            For i = 2 To .Rows
                .Row = i - 1
                If CLng(.Text) = lngTimerID Then
                    .RemoveItem i
                    Exit For
                End If
            Next i
        End If
    End With
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdKillTimer_Click()
    Dim lngTimerID As Long
    msfgTimers.Col = 2
    lngTimerID = CLng(msfgTimers.Text)
    m_CFireTimer.KillTimer lngTimerID
    RemoveTimerFromGrid lngTimerID
End Sub

Private Sub cmdAddTimer_Click()
AddTimerToGrid 0

End Sub

Private Sub cmdSetTimers_Click()
Dim i As Long
Dim strLink As String
Dim lngTime As Long
Dim lngTimerID As Long
Me.Cls

With msfgTimers
    For i = 2 To .Rows
    .Row = i - 1
    .Col = 0
    strLink = .Text
    .Col = 1
    lngTime = CLng(.Text)
    lngTimerID = m_CFireTimer.SetNewTimer(Me, strLink, lngTime)
If lngTimerID Then
    .Col = 2
    .Text = CStr(lngTimerID)
End If
Next i
End With

End Sub

Private Sub Form_Load()

cboLink.AddItem "Do Message"
cboLink.AddItem "Do Print"

msfgTimers.Width = 5412
msfgTimers.ColWidth(0) = 1790
msfgTimers.ColWidth(1) = 1790
msfgTimers.ColWidth(2) = 1790
msfgTimers.Row = 0
msfgTimers.Col = 0


msfgTimers.Text = "Timer Link"
msfgTimers.Col = 1
msfgTimers.Text = "Time"
msfgTimers.Col = 2
msfgTimers.Text = "TimerID"


Set m_CFireTimer = New LinkedTimers.CFireTimer
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set m_CFireTimer = Nothing
End Sub

Private Sub ILinkTimer_FireTimer(ByVal strTimerLink As String, ByVal lngTimerID As Long)
    
    If m_lngTimerID Then
        Me.Print "Timer ID : " & m_lngTimerID & " will reenter function after entry by: " & lngTimerID
        ExecuteTimer strTimerLink, lngTimerID
    Else
        m_lngTimerID = lngTimerID
        ExecuteTimer strTimerLink, m_lngTimerID
        m_lngTimerID = 0
    End If
    RemoveTimerFromGrid lngTimerID
End Sub

Private Sub ExecuteTimer(ByRef strTimerLink As String, ByVal lngTimerID As Long)
    Select Case strTimerLink
    
    Case "Do Message"
        DoMessage lngTimerID
        
    Case "Do Print"
        DoPrint lngTimerID
    End Select
End Sub

Private Sub DoMessage(ByVal lngTimerID As Long)
MsgBox "DoMessage" & "  " & CStr(lngTimerID)
End Sub

Private Sub DoPrint(ByVal lngTimerID As Long)
Me.Print "DoPrint" & "  " & CStr(lngTimerID)
End Sub

