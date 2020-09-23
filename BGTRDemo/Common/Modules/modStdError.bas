Attribute VB_Name = "modStdError"
'********************************************************************************************
'   modStdError Standard Module Definition
'
'   This module defines structures and functions used to temporarily store
'   and subsequently recover error info
'
'   All error handlers are set to Resume Next - if we have problems
'   here we are out of memory, or just have big problems in general!
'   In either event, at this juncture we'll punt!
'********************************************************************************************
Option Explicit

Public Type Error_type
        Number As Long
        Source As String * 255
        Description As String * 255
End Type

'Calculation of string length (in number of characters)
    'Number - long =                                           4 bytes
    'Source - string * 255 char                         510 bytes
    'Description - string * 255 char =                510 bytes
    'Padding for longword alignment =                  0 Bytes
    'Total =                                                   1024 Bytes
    'String length (2 bytes per Unicode char) =  512 Bytes
Public Type ErrorString_type
        ErrorString As String * 512
End Type

'Dim As New so this collection always exists
Private m_colErrorObjects As New Collection

Public Function SaveError(ByVal lngNumber As Long, ByRef strSource As String, ByRef strDescription As String) As String
On Error Resume Next
Dim tmpError As Error_type
Dim tmpErrString As ErrorString_type
Dim strErrorID As String
    tmpError.Number = lngNumber
    tmpError.Source = strSource
    tmpError.Description = strDescription
    
    LSet tmpErrString = tmpError
    strErrorID = CStr(Timer)
    m_colErrorObjects.Add tmpErrString.ErrorString, strErrorID
    
    SaveError = strErrorID
End Function

Public Function ReturnError(ByRef strErrorID As String) As String
On Error Resume Next
    ReturnError = m_colErrorObjects.Item(strErrorID)
    m_colErrorObjects.Remove strErrorID
End Function

Public Sub RaiseError(ByRef strErrorID As String)
Dim tmpErrString As ErrorString_type
Dim tmpError As Error_type

    tmpErrString.ErrorString = m_colErrorObjects.Item(strErrorID)
    m_colErrorObjects.Remove strErrorID
    LSet tmpError = tmpErrString

    Err.Raise tmpError.Number, Trim$(tmpError.Source), Trim$(tmpError.Description)
End Sub

