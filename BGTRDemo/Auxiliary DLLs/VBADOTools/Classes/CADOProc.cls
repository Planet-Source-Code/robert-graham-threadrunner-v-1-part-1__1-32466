VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CADOProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CADOProc Class Definition
'   Class defines methods for handling stored procedures, commands,
'   parameters, and associated objects using VB/ADO
'
'   Instancing is set to:  5 - MultiUse
'
'   Based upon design obtained from:  VBPJ, August 2000.
'
'
'********************************************************************************************
Option Explicit

Public Enum QueryType_enum
    querytype_SQLString = 0
    querytype_StoredProc = 1
End Enum

Public Enum UsesParameters_enum
    uprm_NotUsesParameters = 0
    uprm_UsesParameters = 1
End Enum

Public Enum ReturnsRecords_enum
    retrs_NotReturnsRecords = 0
    retrs_ReturnsRecords = 1
End Enum

Public Enum RecordsetEventsMode_enum
    rsevents_NoEvents = 0
    rsevents_RaiseEvents = 1
End Enum

Public Enum ADOParameterArray_enum
    paramcol_Name = 0
    paramcol_Type = 1
    paramcol_Direction = 2
    paramcol_Size = 3
    paramcol_Value = 4
End Enum

Private m_Parameters() As Variant
Private m_lngParameterCount As Long
Private m_Parameter As ADODB.Parameter
Private m_Cmd As ADODB.Command

Private m_lngCmdTimeOut As Long
Private WithEvents tmpTarget As ADODB.Recordset
Attribute tmpTarget.VB_VarHelpID = -1


'Properties to Get/Let the Command TimeOut value.
Public Property Let CommandTimeOut(lngCmdTimeOut As Long)
   m_lngCmdTimeOut = lngCmdTimeOut
End Property
Public Property Get CommandTimeOut() As Long
   CommandTimeOut = m_lngCmdTimeOut
End Property

'Property to set the number of parameters.
Public Property Let ParameterCount(PrmCount As Long)
   ReDim m_Parameters(PrmCount - 1, paramcol_Name To paramcol_Value)
   'Initially set to zero - will be incremented as parameters are added in SetParameter
   m_lngParameterCount = 0
End Property

'Property to return the value of a parameter
Public Property Get ParameterValue(strKey As String) As Variant
   ParameterValue = m_Cmd.Parameters.Item(strKey)
End Property

'Sub to set the parameter attributes.
Public Sub SetParameter(Optional sName As String = "" _
                       , Optional eType As DataTypeEnum = adEmpty _
                       , Optional eDirection As ParameterDirectionEnum = adParamInput _
                       , Optional lSize As Long _
                       , Optional vValue As Variant _
                       )
On Error GoTo CatchErr

   m_Parameters(m_lngParameterCount, paramcol_Name) = sName
   m_Parameters(m_lngParameterCount, paramcol_Type) = eType
   m_Parameters(m_lngParameterCount, paramcol_Direction) = eDirection
   m_Parameters(m_lngParameterCount, paramcol_Size) = lSize
   m_Parameters(m_lngParameterCount, paramcol_Value) = vValue
   m_lngParameterCount = m_lngParameterCount + 1
Exit Sub
CatchErr:
   Err.Raise Err.Number, Err.Source & " in VBADOTools.CADOProc.SetParameter", Err.Description
End Sub

'Function to execute an adCmdText command instead of a stored proc.
'Return is a no-op - returns -1

Public Sub ExecuteCommand(strCmd As String _
                                , cnConnection As Connection _
                                , Optional lngUseParameters As UsesParameters_enum = uprm_NotUsesParameters _
                                , Optional lngReturnRecordSet As ReturnsRecords_enum = retrs_NotReturnsRecords _
                                , Optional rs As ADODB.Recordset _
                                , Optional lngEvents As RecordsetEventsMode_enum = rsevents_NoEvents _
                                )
On Error GoTo CatchErr
Dim lngI As Long


   Set m_Cmd = New Command
   m_Cmd.CommandType = adCmdText
   m_Cmd.CommandText = strCmd
   m_Cmd.CommandTimeOut = m_lngCmdTimeOut
   
   Set m_Cmd.ActiveConnection = cnConnection
      
   If lngUseParameters = uprm_UsesParameters Then
      For lngI = LBound(m_Parameters, 1) To UBound(m_Parameters, 1)
         Set m_Parameter = m_Cmd.CreateParameter(m_Parameters(lngI, paramcol_Name) _
                                                , m_Parameters(lngI, paramcol_Type) _
                                                , m_Parameters(lngI, paramcol_Direction) _
                                                , m_Parameters(lngI, paramcol_Size) _
                                                , m_Parameters(lngI, paramcol_Value) _
                                                )
         m_Cmd.Parameters.Append m_Parameter
                 
      Next lngI
   End If
      
    If lngReturnRecordSet = retrs_ReturnsRecords Then
        If lngEvents = rsevents_RaiseEvents Then
            rs.Open m_Cmd, , , , adAsyncFetch
        Else
            rs.Open m_Cmd
        End If
   Else
      m_Cmd.Execute , , adExecuteNoRecords
   End If
   Exit Sub
   
CatchErr:
   Err.Raise Err.Number, Err.Source & " in VBADOTools.CADOProc.ExecuteCommand", Err.Description
End Sub


Public Sub ExecuteProcedure(ByVal strProcName As String _
                                , cnConnection As Connection _
                                , Optional lngUseParameters As UsesParameters_enum = uprm_NotUsesParameters _
                                , Optional lngReturnRecordSet As ReturnsRecords_enum = retrs_NotReturnsRecords _
                                , Optional rs As ADODB.Recordset _
                                , Optional lngEvents As RecordsetEventsMode_enum = rsevents_NoEvents _
                                )
On Error GoTo CatchErr
Dim lngI As Long
   
   Set m_Cmd = New Command
   m_Cmd.CommandType = adCmdStoredProc
   m_Cmd.CommandText = strProcName
   m_Cmd.CommandTimeOut = m_lngCmdTimeOut
   
   
   Set m_Cmd.ActiveConnection = cnConnection
   
   If lngUseParameters = uprm_UsesParameters Then
      For lngI = LBound(m_Parameters, 1) To UBound(m_Parameters, 1)
         Set m_Parameter = m_Cmd.CreateParameter(m_Parameters(lngI, paramcol_Name) _
                                                , m_Parameters(lngI, paramcol_Type) _
                                                , m_Parameters(lngI, paramcol_Direction) _
                                                , m_Parameters(lngI, paramcol_Size) _
                                                , m_Parameters(lngI, paramcol_Value) _
                                                )
         m_Cmd.Parameters.Append m_Parameter
                 
      Next lngI
   End If
   
   If lngReturnRecordSet = retrs_ReturnsRecords Then
        If lngEvents = rsevents_RaiseEvents Then
            rs.Open m_Cmd, , , , adAsyncFetch
        Else
            rs.Open m_Cmd
        End If
   Else
      m_Cmd.Execute , , adExecuteNoRecords
   End If
   
   
Exit Sub
CatchErr:
   Err.Raise Err.Number, Err.Source & " in VBADOTools.CADOProc.ExecuteProcedure", Err.Description
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Class Constructor/Destructor
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub Class_Initialize()
    'Set the command timeout to a default of 2 minutes.
    m_lngCmdTimeOut = 120
End Sub

Private Sub Class_Terminate()
   Set m_Parameter = Nothing
   Set m_Cmd = Nothing
End Sub
