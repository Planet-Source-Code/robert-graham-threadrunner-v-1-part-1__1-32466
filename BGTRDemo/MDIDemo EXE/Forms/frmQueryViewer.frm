VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmQueryViewer 
   Caption         =   "Query Viewer"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   7635
   Begin VB.TextBox txtQuery 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Query which generated the current data"
      Top             =   4560
      Width           =   5295
   End
   Begin MSDataGridLib.DataGrid dgQueryView 
      Bindings        =   "frmQueryViewer.frx":0000
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Query results are displayed in grid"
      Top             =   600
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   3
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   4560
      Width           =   975
   End
   Begin MSAdodcLib.Adodc adoQueryView 
      Height          =   330
      Left            =   5160
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoQueryView "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer tmrBlink 
      Left            =   120
      Top             =   0
   End
   Begin VB.Label lblQueryView 
      Caption         =   "Query Results are Read Only!"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label lblQuery 
      Caption         =   "Query:"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   600
   End
   Begin VB.Shape rectNoGrid 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   7  'Diagonal Cross
      Height          =   3735
      Left            =   120
      Top             =   600
      Width           =   7455
   End
End
Attribute VB_Name = "frmQueryViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************************
'   frmQueryViewer Form Definition
'
'   This class defines the QueryViewer form which is instantiated by the
'   QueryRunner - one QueryViewer for each unique query executed.
'   Query results are returned to, and displayed by, this form
'
'********************************************************************************************
Option Explicit

'Form size constants
Private Const DefaultHeight As Long = 5625
Private Const DefaultWidth As Long = 7635

'Factors used in resizing grid
Private Const TOP_OFFSET As Long = 240
Private Const CLOSE_OFFSET As Long = 200
Private Const WIDTH_FACTOR As Double = 0.935

Public strTaskID As String
Private rsResults As ADODB.Recordset

'Public variable for returns records
Public bIsReturnsRecords As Boolean

'Cache the caret blink rate
Private lngBlinkRate As Long
'Cache last state
Private lngLastBlinkState As Long
'Cache QueryRunner deactivation
Private bIsQueryRunnerDeactivate As Boolean


'Public sub called by QueryRunner to clear existing results
'if a new query is run
Public Sub ClearResults()
    On Error Resume Next
    rsResults.Close
    Set rsResults = Nothing
    ShowGrid False
End Sub

'This is only called by QueryRunner

Public Sub DeactivateViewer()
    bIsQueryRunnerDeactivate = True
    Form_Deactivate
End Sub
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ 'This is the function required by the Dispatcher to
'^  return results from task execution
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Public Sub ReturnTaskResults(ByRef arResultsMemento() As Byte, ByVal lngStatus As Long)
On Error GoTo CatchErr

    Select Case lngStatus
        
        Case status_Canceled
            ShowGrid False
            Me.lblQueryView.Caption = "Query was canceled."
        
        Case status_Completed
            If bIsReturnsRecords Then
                ShowGrid True
                ProcessTaskResults arResultsMemento
            Else
                ShowGrid False
                Me.lblQueryView.Caption = "Query completed successfully!    The query does not return records."
            End If
            
        Case status_Aborted
            ShowGrid False
            Me.lblQueryView.Caption = "Query terminated due to errors!"
            RaiseTaskError arResultsMemento
            
    End Select
    
    
 Exit Sub
CatchErr:
    MsgBox "Error Number: " & Err.Number & " Source: " & Err.Source & " raised from frmQueryViewer.ReturnTaskResults" & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

'Function whichg processes the task results and populates the grid

Private Sub ProcessTaskResults(ByRef arResultsMemento() As Byte)
On Error GoTo CatchErr
Dim rstrm As New ADODB.Stream
Dim tmpString As String

    ClearResults
    
    Set rsResults = New ADODB.Recordset
    rstrm.Type = adTypeBinary
    rstrm.Mode = adModeReadWrite
    rstrm.Open
    'Write the results to the stream
    rstrm.Write arResultsMemento
rstrm.SetEOS
rstrm.Position = 0
With rsResults
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open rstrm, , adOpenDynamic, adLockOptimistic
End With

    rstrm.Close
    Set rstrm = Nothing
    Set adoQueryView.Recordset = rsResults
    lblQueryView.Caption = "Query results are read only!"
    AdjustColWidths
    ShowGrid True
    
    CheckActiveStatus
    
Exit Sub
CatchErr:
    MsgBox Err.Number & "  frmQueryViewer.ReturnTaskResults" & vbCrLf & Err.Description
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
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
 If tmrBlink.Interval > 0 Then
    tmrBlink.Interval = 0
    Debug.Print lngLastBlinkState
    lngLastBlinkState = FlashWindow(Me.hwnd, INVERT)
    Debug.Print lngLastBlinkState
 
    If lngLastBlinkState = 0 Then
        lngLastBlinkState = FlashWindow(Me.hwnd, INVERT)
        Debug.Print "In IF:" & lngLastBlinkState
    End If
End If

End Sub

Private Sub Form_Deactivate()
    'If called from QueryRunner, set focus back
    If bIsQueryRunnerDeactivate Then
        bIsQueryRunnerDeactivate = False
        frmQueryRunner.SetFocus
    End If
End Sub

Private Sub Form_Resize()
    ResizeQueryView
End Sub

'This function resizes the grid dimensions to fit the form
'after a resize event.  The numbers are not purely math, and
'include some factors to make the overall appearance more even

Private Sub ResizeQueryView()
On Error Resume Next
Dim lngTotalTopOffset As Long
Dim lngTotalBottomOffset As Long

    lngTotalTopOffset = Me.lblQueryView.Height + 2 * TOP_OFFSET
    lngTotalBottomOffset = Me.cmdClose.Height + 3 * CLOSE_OFFSET + TOP_OFFSET


    Me.dgQueryView.Width = Me.Width * 0.96
    Me.dgQueryView.Height = Me.Height - (lngTotalTopOffset + lngTotalBottomOffset)

    Me.dgQueryView.Left = (Me.Width - Me.dgQueryView.Width) / 2
    Me.dgQueryView.Left = Me.dgQueryView.Left * 0.7
    Me.dgQueryView.Top = Me.lblQueryView.Top + Me.lblQueryView.Height + TOP_OFFSET

    Me.rectNoGrid.Width = Me.dgQueryView.Width
    Me.rectNoGrid.Height = Me.dgQueryView.Height
    
    Me.rectNoGrid.Left = dgQueryView.Left
    Me.rectNoGrid.Top = dgQueryView.Top
    
    Me.cmdClose.Top = Me.dgQueryView.Top + Me.dgQueryView.Height + CLOSE_OFFSET
    Me.cmdClose.Left = Me.dgQueryView.Left + Me.dgQueryView.Width - Me.cmdClose.Width

    Me.lblQuery.Top = Me.cmdClose.Top
    Me.txtQuery.Top = Me.cmdClose.Top
    
    Me.lblQuery.Left = Me.lblQueryView.Left
    Me.txtQuery.Left = Me.lblQuery.Left + lblQuery.Width
    
    If Me.cmdClose.Left < (Me.txtQuery.Left + Me.txtQuery.Width) Then
        Me.txtQuery.Visible = False
        Me.lblQuery.Visible = False
    Else
        Me.txtQuery.Visible = True
        Me.lblQuery.Visible = True
    End If
    
    AdjustColWidths

End Sub

'This function resizes the grid column widths to fit columns to total
'grid width.  The numbers are not purely math, and
'include some factors to make the overall appearance more even

Private Sub AdjustColWidths()
On Error Resume Next
Dim i As Long
Dim count As Long
Dim TotalWidth As Double
Dim dblFactor As Double

    With Me.dgQueryView
        count = .Columns.count
        For i = 0 To count - 1
            TotalWidth = TotalWidth + .Columns.Item(i).Width
        Next i
            
            dblFactor = (.Width / TotalWidth) * WIDTH_FACTOR
    
        
        For i = 0 To count - 1
            .Columns.Item(i).Width = dblFactor * (.Columns.Item(i).Width)
        Next i
    End With
    
End Sub
    
Private Sub ShowGrid(ByVal bIsShow As Boolean)
    
    Me.dgQueryView.Visible = bIsShow
    Me.rectNoGrid.Visible = CBool(1 + CLng(bIsShow))
    
End Sub

   
Private Sub Form_Load()
    g_FormsCount = g_FormsCount + 1
    lngBlinkRate = GetCaretBlinkTime
End Sub

Private Sub Form_Unload(Cancel As Integer)
    g_FormsCount = g_FormsCount - 1
End Sub



Private Sub tmrBlink_Timer()
lngLastBlinkState = FlashWindow(Me.hwnd, INVERT)
End Sub
