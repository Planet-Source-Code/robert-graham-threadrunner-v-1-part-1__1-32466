VERSION 5.00
Begin VB.Form frmSPValues 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stored Procedure Parameter Entry"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   4550
      TabIndex        =   8
      ToolTipText     =   "Click to update a parameter being edited"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   3295
      TabIndex        =   7
      ToolTipText     =   "Click to remove selected parameter"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "Click to add parameter"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5805
      TabIndex        =   9
      ToolTipText     =   "Click to save parameters and close"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters"
      Height          =   2655
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   2895
      Begin VB.ListBox lstParameters 
         Height          =   2010
         Left            =   240
         TabIndex        =   0
         ToolTipText     =   "Click to select a parameter to edit or remove"
         Top             =   400
         Width           =   2415
      End
   End
   Begin VB.Frame fraQuery 
      Caption         =   "Parameter Attributes"
      Height          =   2655
      Left            =   3240
      TabIndex        =   10
      Top             =   240
      Width           =   3660
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "Enter a parameter name"
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtSize 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         ToolTipText     =   $"frmSPValues.frx":0000
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cboADODirection 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         ToolTipText     =   "Select parameter direction"
         Top             =   1695
         Width           =   2295
      End
      Begin VB.ComboBox cboADOType 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         ToolTipText     =   "Select parameter data type"
         Top             =   1230
         Width           =   2295
      End
      Begin VB.TextBox txtValue 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         ToolTipText     =   "Enter a parameter value"
         Top             =   795
         Width           =   2295
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblSize 
         Alignment       =   1  'Right Justify
         Caption         =   "Size:"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblDirection 
         Alignment       =   1  'Right Justify
         Caption         =   "Direction:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "Type:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label lblParameters 
         Alignment       =   1  'Right Justify
         Caption         =   "Value:"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   810
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSPValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************************
'   frmSPValues Form Definition
'
'   This class defines the Parameter Value form which is a modal form
'   form used to input query parameters.
'   This form is displayed upon request by the QueryRunner form.
'
'********************************************************************************************
Option Explicit


Private lngCurrentParameter As Long
Private arParameters() As Variant
Private m_ActiveQuery As BGTRDBWorks.CBGTRUserQuery

Private lngCount As Long
Private arTempParameters() As Variant

Private Sub Form_Load()
    'Populate combos
    LoadADOTypes
    LoadADODirection
    'Set the local active query class to the same instance as frmQueryRunner
    SetCurrentQuery frmQueryRunner.m_ActiveQuery
End Sub

'This function synchronizes the UserQuery class instance with the active instance
'in the QueryRunner form, and loads any exisyting parameters
Public Sub SetCurrentQuery(ByRef oActiveQuery As BGTRDBWorks.CBGTRUserQuery)
Dim i As Long

    Set m_ActiveQuery = oActiveQuery
    arParameters = m_ActiveQuery.ADOParameters
    lngCount = GetParameterCount
    Me.lstParameters.Clear

    For i = 0 To lngCount - 1
        Me.lstParameters.AddItem arParameters(i, paramcol_Name)
    Next i
End Sub

Private Function GetParameterCount() As Long
On Error GoTo CatchErr
    GetParameterCount = UBound(arParameters) + 1
Exit Function
CatchErr:
    GetParameterCount = 0
End Function

Private Sub UpdateFormFields(ByVal lngIndex As Long)
    Me.txtName = arParameters(lngIndex, paramcol_Name)
    Me.txtValue = arParameters(lngIndex, paramcol_Value)
    Me.cboADOType = GetADOType(arParameters(lngIndex, paramcol_Type))
    Me.cboADODirection.ListIndex = arParameters(lngIndex, paramcol_Direction)
    Me.txtSize = arParameters(lngIndex, paramcol_Size)
End Sub

'Add will create a new parameter entry and add
'it to the parameter array for the active query
Private Sub cmdAdd_Click()
Dim i As Long
Dim j As Long

lngCount = lngCount + 1

ReDim arTempParameters(0)
ReDim arTempParameters(0 To lngCount - 1, paramcol_Name To paramcol_Value)

    For i = 0 To lngCount - 2
        For j = paramcol_Name To paramcol_Value
            arTempParameters(i, j) = arParameters(i, j)
        Next j
    Next i
    
    arTempParameters(lngCount - 1, paramcol_Name) = txtName
    arTempParameters(lngCount - 1, paramcol_Value) = txtValue
    arTempParameters(lngCount - 1, paramcol_Type) = SetADOType(cboADOType)
    arTempParameters(lngCount - 1, paramcol_Direction) = cboADODirection.ListIndex
    arTempParameters(lngCount - 1, paramcol_Size) = txtSize
    
    
    Me.lstParameters.AddItem txtName
    arParameters = arTempParameters
    
End Sub


Private Sub cmdRemove_Click()
Dim i As Long
Dim j As Long
Dim k As Long


ReDim arTempParameters(0)
    If lngCount - 1 > 0 Then
        ReDim arTempParameters(0 To lngCount - 2, paramcol_Name To paramcol_Value)
        k = 0
    
        For i = 0 To lngCount - 1
            If arParameters(i, paramcol_Name) <> lstParameters Then
                For j = paramcol_Name To paramcol_Value
                    arTempParameters(k, j) = arParameters(i, j)
                Next j
            k = k + 1
            End If
        Next i
    End If
    
    arParameters = arTempParameters
    lngCount = lngCount - 1
    lstParameters.RemoveItem (lstParameters.ListIndex)
    ClearFields
    
End Sub

Private Sub cmdUpdate_Click()
Dim i As Long
Dim j As Long

    For i = 0 To lngCount - 1
        If arParameters(i, paramcol_Name) = lstParameters Then
            arParameters(i, paramcol_Name) = txtName
            arParameters(i, paramcol_Value) = txtValue
            arParameters(i, paramcol_Type) = SetADOType(cboADOType)
            arParameters(i, paramcol_Direction) = cboADODirection.ListIndex
            arParameters(i, paramcol_Size) = txtSize
            Exit For
        End If
    Next i
        
    Me.lstParameters.Clear
    For i = 0 To lngCount - 1
        Me.lstParameters.AddItem arParameters(i, paramcol_Name)
    Next i
    
End Sub

Private Sub lstParameters_Click()
    lngCurrentParameter = lstParameters.ListIndex
    UpdateFormFields lngCurrentParameter
End Sub

Private Sub cmdClose_Click()
    
    'Update the active UserQuery class instance
    m_ActiveQuery.ADOParameters = arParameters
    m_ActiveQuery.ParameterCount = GetParameterCount
    Set m_ActiveQuery = Nothing
    Unload Me
    
End Sub


'Type conversion functions for ADO type definitions
'The enumerated constants do not adhere to contiguous order
'so this must be performed manually
Private Function SetADOType(ByRef strADOType As String) As Long

    Select Case strADOType
    
        Case "adBinary"
            SetADOType = 128
        Case "adBoolean"
            SetADOType = 11
        Case "adBSTR"
            SetADOType = 8
        Case "adChar"
            SetADOType = 129
        Case "adCurrency"
            SetADOType = 6
        Case "adDate"
            SetADOType = 7
        Case "adDBTimeStamp"
            SetADOType = 135
        Case "adDecimal"
            SetADOType = 14
        Case "adDouble"
            SetADOType = 5
        Case "adInteger"
            SetADOType = 3
        Case "adNumeric"
            SetADOType = 131
        Case "adSingle"
            SetADOType = 4
        Case "adSmallInt"
            SetADOType = 2
        Case "adUnsignedInt"
            SetADOType = 19
        Case "adUserDefined"
            SetADOType = 132
        Case "adVarChar"
            SetADOType = 200
        Case "adVariant"
            SetADOType = 12
        Case Else
            SetADOType = 12
    End Select
End Function

Private Function GetADOType(ByVal lngADOType As Long) As String

    Select Case lngADOType
    
        Case 128
            GetADOType = "adBinary"
        Case 11
            GetADOType = "adBoolean"
        Case 8
            GetADOType = "adBSTR"
        Case 129
            GetADOType = "adChar"
        Case 6
            GetADOType = "adCurrency"
        Case 7
            GetADOType = "adDate"
        Case 135
            GetADOType = "adDBTimeStamp"
        Case 14
            GetADOType = "adDecimal"
        Case 5
            GetADOType = "adDouble"
        Case 3
            GetADOType = "adInteger"
        Case 131
            GetADOType = "adNumeric"
        Case 4
            GetADOType = "adSingle"
        Case 2
            GetADOType = "adSmallInt"
        Case 19
            GetADOType = "adUnsignedInt"
        Case 132
            GetADOType = "adUserDefined"
        Case 200
            GetADOType = "adVarChar"
        Case 12
            GetADOType = "adVariant"
        Case Else
            GetADOType = "adVariant"
    End Select
End Function

Private Sub LoadADOTypes()
    With Me.cboADOType
        .Clear
        .AddItem "adBinary"
        .AddItem "adBoolean"
        .AddItem "adBSTR"
        .AddItem "adChar"
        .AddItem "adCurrency"
        .AddItem "adDate"
        .AddItem "adDBTimeStamp"
        .AddItem "adDecimal"
        .AddItem "adDouble"
        .AddItem "adInteger"
        .AddItem "adNumeric"
        .AddItem "adSingle"
        .AddItem "adSmallInt"
        .AddItem "adUnsignedInt"
        .AddItem "adUserDefined"
        .AddItem "adVarChar"
        .AddItem "adVariant"
    End With
End Sub

Private Sub LoadADODirection()
    With Me.cboADODirection
        .Clear
        .AddItem "adParamUnknown"
        .AddItem "adParamInput"
        .AddItem "adParamOutput"
        .AddItem "adParamInputOutput"
        .AddItem "adParamReturnValue"
    End With
End Sub

Private Sub ClearFields()

    Me.txtName = vbNullString
    Me.txtSize = vbNullString
    Me.txtValue = vbNullString
    
    Me.cboADODirection.ListIndex = -1
    Me.cboADOType.ListIndex = -1
    
End Sub
