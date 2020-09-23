VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTaskViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Task Viewer"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   6435
   Begin VB.TextBox txtSelected 
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelTask 
      Caption         =   "Cancel Task"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvTaskViewer 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgTaskViewer"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "TaskID"
         Object.Tag             =   "TaskID"
         Text            =   "TaskID"
         Object.Width           =   2522
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "TaskStatus"
         Object.Tag             =   "TaskStatus"
         Text            =   "Task Status"
         Object.Width           =   2522
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "TaskDescription"
         Object.Tag             =   "TaskDescription"
         Text            =   "Task Description"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTaskViewer 
      Left            =   480
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskViewer.frx":0000
            Key             =   "BGTRTask"
            Object.Tag             =   "BGTRTask"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblSelected 
      Caption         =   "Selected Task:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmTaskViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefaultHeight As Long = 3720
Private Const DefaultWidth As Long = 6525

Public Enum ListColumns_enum
    listcol_Task = 0
    listcol_TaskStatus = 1
    listcol_TaskDescription = 2
End Enum

Private oSelectedItem As ListItem

Public Sub UpdateTask(ByRef strTaskID As String, ByRef strDescription As String, ByVal lngStatus As SubmittedTaskStatus_enum)
Dim oListItem As ListItem

    Select Case lngStatus
    
        Case submitstatus_Active
            With Me.lvTaskViewer
            Set oListItem = .FindItem(strTaskID, lvwText)
            If Not oListItem Is Nothing Then
                oListItem.SubItems(listcol_TaskStatus) = "Active"
            Else
                .ListItems.Add , , strTaskID, , 1
                .ListItems(.ListItems.count).SubItems(listcol_TaskStatus) = "Active"
               .ListItems(.ListItems.count).SubItems(listcol_TaskDescription) = strDescription
            End If
            End With
            
                
        Case submitstatus_Queued
        With Me.lvTaskViewer
            Set oListItem = .FindItem(strTaskID, lvwText)
            If Not oListItem Is Nothing Then
                oListItem.SubItems(listcol_TaskStatus) = "Queued"
            Else
               .ListItems.Add , , strTaskID, , 1
                .ListItems(.ListItems.count).SubItems(listcol_TaskStatus) = "Queued"
               .ListItems(.ListItems.count).SubItems(listcol_TaskDescription) = strDescription
            End If
            End With
        Case submitstatus_Terminated
            With Me.lvTaskViewer
            Set oListItem = .FindItem(strTaskID, lvwText)
            .ListItems.Remove (oListItem.Index)
            
            If Not oSelectedItem Is Nothing Then
                If oSelectedItem = oListItem Then
                    Set oSelectedItem = Nothing
                    txtSelected = vbNullString
                End If
            End If
            
            Set oListItem = Nothing
            
            End With
           Case submitstatus_Canceling
        With Me.lvTaskViewer
            Set oListItem = .FindItem(strTaskID, lvwText)
            If Not oListItem Is Nothing Then
                oListItem.SubItems(listcol_TaskStatus) = "Canceling"
            Else
                .ListItems.Add , , strTaskID, , 1
                .ListItems(.ListItems.count).SubItems(listcol_TaskStatus) = "Canceling"
               .ListItems(.ListItems.count).SubItems(listcol_TaskDescription) = strDescription
            End If
            End With
    End Select

    If oSelectedItem Is Nothing Then
        With lvTaskViewer
            If .ListItems.count > 0 Then
                Set oSelectedItem = .ListItems.Item(FIRST_ITEM)
                txtSelected = oSelectedItem
            End If
        End With
    End If
    
        
End Sub

Private Sub cmdCancelTask_Click()
    
    If Not oSelectedItem Is Nothing Then
        g_CTaskDispatcher.CancelTask oSelectedItem.Text
        UpdateTask oSelectedItem.Text, vbNullString, submitstatus_Canceling
    End If
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Width = DefaultWidth
    Me.Height = DefaultHeight
    CenterForm Me
    g_FormsCount = g_FormsCount + 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Select Case UnloadMode
        Case vbFormCode, _
                vbFormControlMenu, _
                vbFormOwner:
                    Cancel = True
                    Me.Hide
        
        Case Else
                'Do nothing
                
    End Select
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    g_FormsCount = g_FormsCount - 1
End Sub

Private Sub lvTaskViewer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set oSelectedItem = Item
    Me.txtSelected = Item
End Sub
