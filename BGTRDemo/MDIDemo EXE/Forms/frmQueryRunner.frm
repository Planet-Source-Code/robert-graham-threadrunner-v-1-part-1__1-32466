VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQueryRunner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Query Runner"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   7230
   Begin VB.CheckBox chkRecords 
      Caption         =   "Returns Records"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "Check if the query returns records"
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Frame fraCurrentQuery 
      Caption         =   "Current Query"
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   240
      Width           =   6975
      Begin VB.ComboBox cboCurrentQuery 
         Height          =   315
         Left            =   480
         TabIndex        =   0
         ToolTipText     =   "Enter a new query indentifier or select an existing one from the list"
         Top             =   260
         Width           =   6135
      End
   End
   Begin VB.Frame fraQuery 
      Caption         =   "Query Definition"
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Top             =   4440
      Width           =   6975
      Begin VB.TextBox txtQuery 
         Height          =   495
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Enter the SQL statement string - multiple lines may be used"
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label lblQuery 
         Alignment       =   1  'Right Justify
         Caption         =   "SQL Statement:"
         Height          =   255
         Left            =   140
         TabIndex        =   23
         Top             =   360
         Width           =   1390
      End
   End
   Begin VB.Frame fraMode 
      Caption         =   "Query Mode"
      Height          =   735
      Left            =   2880
      TabIndex        =   21
      Top             =   3600
      Width           =   4215
      Begin VB.OptionButton optMode 
         Caption         =   "Stored Procedure"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   9
         ToolTipText     =   "Execute an existing stored procedure"
         Top             =   280
         Width           =   1575
      End
      Begin VB.OptionButton optMode 
         Caption         =   "SQL String"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   8
         ToolTipText     =   "Execute a user defined SQL query"
         Top             =   280
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame fraServer 
      Caption         =   "DB Server Parameters"
      Height          =   2175
      Left            =   2880
      TabIndex        =   15
      Top             =   1200
      Width           =   4215
      Begin VB.CommandButton cmdBrowse 
         Caption         =   ". . ."
         Height          =   285
         Left            =   3600
         TabIndex        =   3
         ToolTipText     =   "Click to select an MDB file"
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   6
         Text            =   "Text1"
         ToolTipText     =   "Enter a password if required"
         Top             =   1700
         Width           =   2295
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "User Name"
         ToolTipText     =   "Enter user name"
         Top             =   1292
         Width           =   2295
      End
      Begin VB.TextBox txtDatabase 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "Database Name"
         ToolTipText     =   "Enter database name"
         Top             =   886
         Width           =   2295
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Text            =   "Server Name"
         ToolTipText     =   "Enter SQL Server Name"
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   330
         Left            =   240
         TabIndex        =   19
         Top             =   1700
         Width           =   855
      End
      Begin VB.Label lblUser 
         Alignment       =   1  'Right Justify
         Caption         =   "UserName:"
         Height          =   330
         Left            =   240
         TabIndex        =   18
         Top             =   1292
         Width           =   855
      End
      Begin VB.Label lblDatabase 
         Alignment       =   1  'Right Justify
         Caption         =   "Database:"
         Height          =   330
         Left            =   240
         TabIndex        =   17
         Top             =   886
         Width           =   855
      End
      Begin VB.Label lblServer 
         Alignment       =   1  'Right Justify
         Caption         =   "Server:"
         Height          =   330
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Top             =   5800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   5800
      Width           =   1215
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Execute"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   5800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSPValues 
      Caption         =   "Parameters"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   5800
      Width           =   1215
   End
   Begin VB.Frame fraDBType 
      Caption         =   "Database Type"
      Height          =   2175
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   2175
      Begin VB.ListBox lstDBType 
         Height          =   1425
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Select a database type"
         Top             =   480
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog cdlgQueryRunner 
      Left            =   0
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmQueryRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************************
'   frmQueryRunner Form Definition
'
'   This class defines the QueryRunner form which is a single instance
'   form used to input query parameters and launch the Query tasks.
'   This form maintains a temporary collection of UserQuery objects.
'   The collection is cleared if the form is closed.
'   This form creates one QueryViewer for each unique query executed.
'   Query results are returned to, and displayed by, the QueryViewer
'
'
'********************************************************************************************
Option Explicit

Private Enum DBType_enum
    dbtype_Unknown = 0
    dbtype_SQLServer = 1
    dbtype_Jet = 2
End Enum

Private strConnection As String
Private strMDBPath As String
Private lngDBServerType As DBType_enum

Private lngQueryType As QueryType_enum
Private lngReturnsRecords As ReturnsRecords_enum

Private oConnection As VBADOTools.CDBConnection
Private arParameters() As Variant

Public m_ActiveQuery As CBGTRUserQuery
Public m_AllQuerys As CBGTRUserQuerys

Private Const DefaultWidth As Long = 7350
Private Const DefaultHeight As Long = 6720

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Query Set Functions - Create/Load query instances
'^ from a local collection maintained as long as form is open
'^ These functions are triggered by cboCurrentQuery events
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub SetSelectedQuery()
Dim i As Long
Dim strQueryName As String
Dim bExists As Boolean

    If Len(cboCurrentQuery) > 0 Then
        'See if query already exists
        strQueryName = cboCurrentQuery
        For i = 0 To cboCurrentQuery.ListCount - 1
            If cboCurrentQuery.List(i) = strQueryName Then
                bExists = True
                Exit For
            End If
        Next i
    
        If Not bExists Then
            'New query - add to the list
            cboCurrentQuery.AddItem strQueryName
        End If
        'Load the query class instance
        SetActiveQuery strQueryName
        
    End If
End Sub

Private Sub SetActiveQuery(ByRef strQuery As String)
    'Get the query instance from the collection
    Set m_ActiveQuery = m_AllQuerys.Item(strQuery)
    If m_ActiveQuery Is Nothing Then
        'New query - create a new class instance
        Set m_ActiveQuery = New CBGTRUserQuery
        m_AllQuerys.Insert m_ActiveQuery, strQuery
    Else
        'Existing query - refresh the form fields
        UpdateFormFields
    End If
    
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Update Functions - update the query class instance from
'^ the form fields, or update form fields from a query class
'^ depending upon the current form action
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub UpdateFormFields()
    With m_ActiveQuery
        Me.txtServer = .Connection.Server
        Me.txtDatabase = .Connection.Database
        Me.txtUser = .Connection.User
        Me.txtPassword = .Connection.Password

        Me.lstDBType.ListIndex = .DBType - 1
        Me.optMode(.QueryType).Value = True
        Me.chkRecords = .ReturnsRecords
        Me.txtQuery.Text = .SQLStatement
    End With
End Sub

Private Sub UpdateQueryClass()
    With m_ActiveQuery
        .Connection.Server = Me.txtServer
        .Connection.Database = Me.txtDatabase
        .Connection.User = Me.txtUser
        .Connection.Password = Me.txtPassword
        .DBType = Me.lstDBType.ListIndex + 1
        .QueryType = lngQueryType
        .ReturnsRecords = chkRecords
       .SQLStatement = Me.txtQuery.Text
    End With
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Query Task functions create a task memento based upon
'^ the currently active query class for submission as a
'^ background task
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Function MakeQueryTask() As Byte()
Dim arFactoryDescriptor() As String
Dim arWorkDescriptor() As String
Dim pbTemp As New PropertyBag
Dim i As Long

ReDim arFactoryDescriptor(0)
arFactoryDescriptor(0) = "DBTasks"

ReDim arWorkDescriptor(0)
arWorkDescriptor(0) = "Query"

    With pbTemp
        .WriteProperty "FactoryDescUB", UBound(arFactoryDescriptor)
        For i = 0 To UBound(arFactoryDescriptor)
        .WriteProperty "FactoryDescItem" & i, arFactoryDescriptor(i)
        Next i
        
        .WriteProperty "WorkDescUB", UBound(arWorkDescriptor)
        For i = 0 To UBound(arWorkDescriptor)
        .WriteProperty "WorkDescItem" & i, arWorkDescriptor(i)
        Next i
        .WriteProperty "WorkMemento", ConfigureQueryTask
        
        MakeQueryTask = .Contents
    End With
End Function

Private Function ConfigureQueryTask() As Byte()
    UpdateQueryClass
    ConfigureQueryTask = m_ActiveQuery.GetSuperState
End Function

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Major form action command buttons
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Display modal form to enter/edit query parameters
Private Sub cmdSPValues_Click()
    frmSPValues.Show vbModal
End Sub

Private Sub cmdRun_Click()
On Error GoTo CatchErr
Dim strTaskID As String
Dim frmQuery As frmQueryViewer
Dim frm As Form
Dim bExists As Boolean

    If ValidateQueryFields Then
        
        'Enumerate foms to see if a QueryViewer instance exists for this query
        'and assign a new one if not.
        
        For Each frm In Forms
            If frm.Caption = "Query Viewer - " & Me.cboCurrentQuery Then
                Set frmQuery = frm
                bExists = True
                Exit For
            End If
        Next frm
    
        If Not bExists Then
            Set frmQuery = New frmQueryViewer
        End If
    
        With frmQuery
            .ClearResults
            .Caption = "Query Viewer - " & Me.cboCurrentQuery
            .lblQueryView.Caption = "Query is running ..."
            'Set the property to return records or not
            .bIsReturnsRecords = lngReturnsRecords
            .txtQuery.Text = Me.txtQuery.Text
            .Show
            .DeactivateViewer
        End With
           
        'Submit task, with a return recipient reference to the asssigned QueryViewer instance
        strTaskID = g_CTaskDispatcher.SubmitTask(MakeQueryTask, frmQuery, "Execute User Query")
    End If
    
Exit Sub
CatchErr:
    MsgBox Err.Number & "  " & Err.Source & " in frmQueryRunner.cmdRun_Click" & vbCrLf & Err.Description, vbExclamation
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    Me.lstDBType.ListIndex = -1
    ClearFields
End Sub

Private Sub cmdClose_Click()
    Unload frmQueryRunner
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Form controls
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub cboCurrentQuery_Click()
    
    If Not m_ActiveQuery Is Nothing Then
        UpdateQueryClass
    End If
    SetSelectedQuery

End Sub

Private Sub cboCurrentQuery_LostFocus()
     
     If Not m_ActiveQuery Is Nothing Then
        UpdateQueryClass
    End If
    SetSelectedQuery

End Sub

Private Sub lstDBType_Click()

    If lstDBType.List(lstDBType.ListIndex) = "Jet Database" Then
        lblServer.Caption = "DB File:"
        cmdBrowse.Visible = True
        txtServer.ToolTipText = "Enter the full path for the MDB file or click browse to select"
        txtDatabase.ToolTipText = "Not required for Jet MDB databases"
    Else
        lblServer.Caption = "Server:"
        cmdBrowse.Visible = False
        txtServer.ToolTipText = "Enter a SQL Server name"
        txtDatabase.ToolTipText = "Enter a database name"
        
    End If

    lngDBServerType = lstDBType.ListIndex + 1
    
End Sub

Private Sub chkRecords_Click()
    
    If chkRecords Then
        lngReturnsRecords = retrs_ReturnsRecords
    Else
        lngReturnsRecords = retrs_NotReturnsRecords
    End If
    
End Sub

Private Sub optMode_Click(Index As Integer)
    
    lngQueryType = Index
    
    Select Case Index
        
        Case querytype_SQLString
            lblQuery.Caption = "SQL Statement:"
            Me.txtQuery.ToolTipText = "Enter the SQL statement string - multiple lines may be used"
            Me.cmdSPValues.Enabled = False
                
        Case querytype_StoredProc
            lblQuery.Caption = "Stored Proc Name:"
            Me.txtQuery.ToolTipText = "Enter the stored procedure name -  - multiple lines may be used"
            Me.cmdSPValues.Enabled = True
    End Select
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo CatchErr
Dim tmpString As String
Dim strDBName As String
Dim strTest As String

    cdlgQueryRunner.Flags = cdlOFNFileMustExist
    cdlgQueryRunner.ShowOpen
    strMDBPath = cdlgQueryRunner.FileName

    tmpString = Left(strMDBPath, (Len(strMDBPath) - 4))
    strTest = ""

    Do While strTest <> "\"
        strTest = Right(tmpString, 1)
        strDBName = strTest & strDBName
        tmpString = Left(tmpString, (Len(tmpString) - 1))
    Loop

    strDBName = Right(strDBName, Len(strDBName) - 1)
    Me.txtServer = strMDBPath
    Me.txtDatabase = strDBName

Exit Sub
CatchErr:
    strMDBPath = ""
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Control support functions
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Function ValidateQueryFields() As Boolean
    
    If Len(Me.cboCurrentQuery) = 0 Then
        MsgBox "Please enter a query identifier in the current query field", vbExclamation
        Exit Function
    End If

    If Len(Me.txtServer) = 0 Then
        MsgBox "Please enter a server name or select a Jet database file", vbExclamation
        Exit Function
    End If
    
    If Len(Me.txtDatabase) = 0 Then
        MsgBox "Please enter a database!", vbExclamation
        Exit Function
    End If
    
    If Len(Me.txtQuery.Text) = 0 Then
        MsgBox "Please enter a query statement or stored procedure name!", vbExclamation
        Exit Function
    End If

    'Made it this far!
    ValidateQueryFields = True
    
End Function

Private Sub InitializeForm()
    ClearFields
    Me.lstDBType.AddItem "SQL Server"
    Me.lstDBType.AddItem "Jet Database"
End Sub

Private Sub ClearFields()
On Error Resume Next
    Me.txtServer = vbNullString
    Me.txtPassword = vbNullString
    Me.txtDatabase = vbNullString
    Me.txtUser = vbNullString
    Me.txtQuery.Text = vbNullString
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Form Class Constructor/Destructor
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub Form_Load()
    Me.Width = DefaultWidth
    Me.Height = DefaultHeight
    CenterForm Me
    g_FormsCount = g_FormsCount + 1
    InitializeForm
    Set m_AllQuerys = New CBGTRUserQuerys
    lngReturnsRecords = retrs_ReturnsRecords
End Sub

Private Sub Form_Unload(Cancel As Integer)
    g_FormsCount = g_FormsCount - 1
End Sub
