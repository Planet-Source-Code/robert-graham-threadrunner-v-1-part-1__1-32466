VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmQueryRunner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Query Runner"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   7230
   Begin VB.Frame fraCurrentQuery 
      Caption         =   "Current Query"
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   240
      Width           =   6975
      Begin VB.ComboBox cboCurrentQuery 
         Height          =   315
         Left            =   480
         TabIndex        =   23
         Top             =   260
         Width           =   6135
      End
   End
   Begin VB.Frame fraQuery 
      Caption         =   "Query Definition"
      Height          =   1095
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Width           =   6975
      Begin RichTextLib.RichTextBox rtxtQuery 
         Height          =   495
         Left            =   1560
         TabIndex        =   21
         ToolTipText     =   "Enter the SQL statement string"
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmStarter.frx":0000
      End
      Begin VB.Label lblQuery 
         Alignment       =   1  'Right Justify
         Caption         =   "SQL Statement:"
         Height          =   255
         Left            =   140
         TabIndex        =   20
         Top             =   360
         Width           =   1390
      End
   End
   Begin VB.Frame fraMode 
      Caption         =   "Query Mode"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   6975
      Begin VB.OptionButton optMode 
         Caption         =   "Stored Procedure"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   18
         ToolTipText     =   "Execute an existing stored procedure"
         Top             =   280
         Width           =   1575
      End
      Begin VB.OptionButton optMode 
         Caption         =   "SQL String"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   17
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
      TabIndex        =   4
      Top             =   1200
      Width           =   4215
      Begin VB.CommandButton cmdBrowse 
         Caption         =   ". . ."
         Height          =   285
         Left            =   3600
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1700
         Width           =   2295
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   "User Name"
         Top             =   1292
         Width           =   2295
      End
      Begin VB.TextBox txtDatabase 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "Database Name"
         Top             =   886
         Width           =   2295
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "Server Name"
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   330
         Left            =   240
         TabIndex        =   12
         Top             =   1700
         Width           =   855
      End
      Begin VB.Label lblUser 
         Alignment       =   1  'Right Justify
         Caption         =   "UserName:"
         Height          =   330
         Left            =   240
         TabIndex        =   10
         Top             =   1292
         Width           =   855
      End
      Begin VB.Label lblDatabase 
         Alignment       =   1  'Right Justify
         Caption         =   "Database:"
         Height          =   330
         Left            =   240
         TabIndex        =   8
         Top             =   886
         Width           =   855
      End
      Begin VB.Label lblServer 
         Alignment       =   1  'Right Justify
         Caption         =   "Server:"
         Height          =   330
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   5740
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   5740
      Width           =   1215
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Execute"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   5740
      Width           =   1215
   End
   Begin VB.CommandButton cmdSPValues 
      Caption         =   "Stored Proc Values"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   5740
      Width           =   1215
   End
   Begin VB.Frame fraDBType 
      Caption         =   "Database Type"
      Height          =   2175
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   2175
      Begin VB.ListBox lstDBType 
         Height          =   1425
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog cdlgQueryRunner 
      Left            =   0
      Top             =   5520
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
OptionExplicit

Private Enum DBType_enum
    dbtype_Unknown = 0
    dbtype_SQLServer = 1
    dbtype_Jet = 2
End Enum

Private bExists As Boolean

Private strConnection As String
Private strMDBPath As String
Private lngDBServerType As DBType_enum

Private lngQueryMode As Long
Private lngQueryType As QueryType_enum
Private lngReturnsRecords As ReturnsRecords_enum

Private oConnection As BGTRDBWorks.CBGTRConnection
Private arParameters() As Variant

Public m_ActiveQuery As CBGTRUserQuery
Public m_AllQuerys As CBGTRUserQuerys


Private Const DefaultWidth As Long = 7350
Private Const DefaultHeight As Long = 6720




'Private Sub MakeJetProcs1()
'Dim objCADOProc As CADOProc
'Dim cnConnection As ADODB.Connection
'Dim strParameter As String
'Dim strType As String
'Dim strCommand As String
'
'Dim arParameters() As Variant
'
'MakeConnectionString
'
'Set cnConnection = New ADODB.Connection
'cnConnection.ConnectionString = strConnection
'
'cnConnection.Open
'Set objCADOProc = New CADOProc
'
'ReDim arParameters(0 To 1, 0 To 1)
'
'    strType = "INTEGER"
'    strParameter = "bDeleteRecords"
'    arParameters(0, 0) = strParameter
'    arParameters(0, 1) = strType
'
'    strType = "INTEGER"
'    strParameter = "lngRowCount"
'    arParameters(1, 0) = strParameter
'    arParameters(1, 1) = strType
'
'    strCommand = "SET ROWCOUNT [Row Count] " & _
'                           "IF [bDeleteRecords] = -1 " & _
'                           "DELETE FROM SMTDemoDBTable1_Temp " & _
'                            "SELECT * From SMTDemoDBTable1_Temp"
'
'    objCADOProc.CreateJetProcedure cnConnection, "sp_Archive_SMTDemoDBTable1_Temp", strCommand, 2, arParameters
'
'    ReDim arParameters(0)
'    ReDim arParameters(0 To 1, 0 To 1)
'
'    strType = "DATETIME"
'    strParameter = "Start Date"
'
'    arParameters(0, 0) = strParameter
'    arParameters(0, 1) = strType
'
'    strParameter = "End Date"
'
'    arParameters(1, 0) = strParameter
'    arParameters(1, 1) = strType
'
'    strCommand = "CREATE TABLE SMTDemoDBTable1_Temp " & _
'                            "(RecordNo INT," & _
'                            "SMTTimeStamp DATETIME," & _
'                              "Field1 INT," & _
'                              "Field2 INT," & _
'                              "Field3 INT) " & _
'                        "INSERT INTO SMTDemoDBTable1_Temp " & _
'                        "SELECT        RecordNo," & _
'                                            "SMTTimeStamp," & _
'                                            "Field1," & _
'                                            "Field2," & _
'                                            "Field3 " & _
'                          "FROM SMTDemoDBTable1 " & _
'                          "WHERE SMTTimeStamp >= [Start Date] " & _
'                          "AND SMTTimeStamp   <= [End Date] " & _
'                          "ORDER BY RecordNo"
'
'    objCADOProc.CreateJetProcedure cnConnection, "sp_Make_SMTDemoDBTable1_Temp", strCommand, 2, arParameters
'
'
'End Sub
'

Private Sub cboCurrentQuery_Click()
    SetSelectedQuery
End Sub

Private Sub cboCurrentQuery_LostFocus()
    SetSelectedQuery
End Sub

Private Sub SetSelectedQuery()
Dim i As Long
Dim strQueryName As String

    If Len(cboCurrentQuery) > 0 Then
        strQueryName = cboCurrentQuery
        For i = 0 To cboCurrentQuery.ListCount - 1
            If cboCurrentQuery.List(i) = strQueryName Then
                bExists = True
                Exit For
            End If
        Next i
    
        If Not bExists Then
            cboCurrentQuery.AddItem strQueryName
        End If
        
        SetActiveQuery strQueryName
        
    End If
End Sub

Private Sub SetActiveQuery(ByRef strQuery As String)
    Set m_ActiveQuery = m_AllQuerys.Item(strQuery)
    If m_ActiveQuery Is Nothing Then
        Set m_ActiveQuery = New CBGTRUserQuery
        m_AllQuerys.Insert m_ActiveQuery, strQuery
    End If
    UpdateFormFields
End Sub

Private Sub UpdateFormFields()
    With m_ActiveQuery
        Me.txtServer = .Connection.Server
        Me.txtDatabase = .Connection.Database
        Me.txtUser = .Connection.User
        Me.txtPassword = .Connection.Password

        Me.lstDBType.ListIndex = .DBType - 1
        Me.optMode(.QueryType).Value = True
        Me.chkRecords = .ReturnsRecords
        Me.rtxtQuery.TextRTF = .SQLStatement
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
        .ReturnsRecords = lngReturnsRecords
       .SQLStatement = Me.rtxtQuery.Text
    End With
End Sub

Private Sub chkRecords_Click()
    If chkRecords Then
        lngReturnsRecords = retrs_ReturnsRecords
    Else
        lngReturnsRecords = retrs_NotReturnsRecords
    End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    Me.lstDBType.ListIndex = -1
    ClearFields
End Sub

Private Sub cmdRun_Click()
Dim strTaskID As String
Dim frmQuery As frmQueryViewer
Dim frm As Form
Dim bExists As Boolean

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
        .Caption = "Query Viewer - " & Me.cboCurrentQuery
        .rtxtQueryResults = "Query is running ..."
        .Show
    End With

    frmMDIMain.txtThread.Visible = True
    Select Case lngDBServerType
        Case dbtype_Jet
            strTaskID = g_CTaskDispatcher.SubmitTask(MakeAccessTask, frmQuery, "Execute Jet Query")
        Case dbtype_SQLServer
            strTaskID = g_CTaskDispatcher.SubmitTask(MakeSQLTask, frmQuery, "Execute SQL Query")
    End Select
    
End Sub

Private Sub cmdSPValues_Click()
    frmSPValues.Show vbModal
End Sub

Private Sub Form_Load()
    Me.Width = DefaultWidth
    Me.Height = DefaultHeight
    CenterForm Me
    g_FormsCount = g_FormsCount + 1
    InitializeForm
    Set m_AllQuerys = New CBGTRUserQuerys
End Sub

Private Sub Form_Unload(Cancel As Integer)
    g_FormsCount = g_FormsCount - 1
End Sub

Private Sub lstDBType_Click()

    If lstDBType.List(lstDBType.ListIndex) = "Jet Database" Then
        lblServer.Caption = "DB File:"
        cmdBrowse.Visible = True
    Else
        lblServer.Caption = "Server:"
        cmdBrowse.Visible = False
    End If

    lngDBServerType = lstDBType.ListIndex + 1
    ClearFields
End Sub

Private Sub optMode_Click(Index As Integer)
    
    lngQueryMode = Index
    lngQueryType = Index
    
    Select Case Index
        
        Case querytype_SQLString
            lblQuery.Caption = "SQL Statement:"
            Me.rtxtQuery.ToolTipText = "Enter the SQL statement string"
            Me.cmdSPValues.Enabled = False
                
        Case querytype_StoredProc
            lblQuery.Caption = "Stored Proc Name:"
            Me.rtxtQuery.ToolTipText = "Enter the stored procedure name"
            Me.cmdSPValues.Enabled = True
            frmSPValues.Show vbModal
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

Private Sub cmdClose_Click()
    Unload frmQueryRunner
End Sub

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
    Me.rtxtQuery.Text = vbNullString
End Sub

Private Function MakeAccessQuery() As Byte()

    Set oConnection = New BGTRDBWorks.CBGTRConnection
    oConnection.Database = Me.txtDatabase
    oConnection.Server = Me.txtServer
    oConnection.User = Me.txtUser
    oConnection.Password = Me.txtPassword
    
    m_ActiveQuery.Connection = oConnection
    
    m_ActiveQuery.SQLStatement = Me.rtxtQuery.Text
    
    
    MakeAccessQuery = m_ActiveQuery.GetSuperState
    
End Function

Private Function MakeAccessTask() As Byte()
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
        .WriteProperty "WorkMemento", MakeAccessQuery
        
        MakeAccessTask = .Contents
    End With
    
End Function

Private Function MakeSQLQuery() As Byte()

    Set oConnection = New BGTRDBWorks.CBGTRConnection
    oConnection.Database = Me.txtDatabase
    oConnection.Server = Me.txtServer
    oConnection.User = Me.txtUser
    oConnection.Password = Me.txtPassword
    
    m_ActiveQuery.Connection = oConnection
    m_ActiveQuery.SQLStatement = Me.rtxtQuery.Text
    m_ActiveQuery.ADOParameters = arParameters
    
    
    MakeSQLQuery = m_ActiveQuery.GetSuperState
    
End Function

Private Function MakeSQLTask() As Byte()
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
        .WriteProperty "WorkMemento", MakeSQLQuery
        
        MakeSQLTask = .Contents
    End With
End Function



