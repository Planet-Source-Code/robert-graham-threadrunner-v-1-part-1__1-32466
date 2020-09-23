Attribute VB_Name = "modFiles"
'********************************************************************************************
'   modFiles Standard Module Definition
'
'   This module defines common file routines
'
'********************************************************************************************
Option Explicit

Public Function ReadFile(strFile As String) As Byte()
On Error GoTo CatchErr
Dim lFile As Long
Dim vStream As Variant
    If FileExists(strFile) Then
    lFile = FreeFile
    Open strFile For Binary As lFile
        Get lFile, , vStream
    Close lFile
    ReadFile = vStream
    Exit Function
    Else
        Err.Raise VB_ERR_FILE_NOT_FOUND
    End If
Exit Function
CatchErr:
    Dim strErrorID As String
    strErrorID = SaveError(Err.Number, "modFiles.ReadFile", Err.Description)
    On Error Resume Next
    Close lFile
    RaiseError strErrorID
End Function

Public Sub WriteFile(ByRef strFile As String, ByRef baFile() As Byte)
On Error GoTo CatchErr
Dim lFile As Long
Dim vStream As Variant
    If Len(strFile) > 0 Then
        vStream = baFile()
        lFile = FreeFile
        Open strFile For Binary As lFile
            Put lFile, , vStream
        Close lFile
    Else
        Err.Raise VB_ERR_BAD_FILE_NAME
    End If
Exit Sub
CatchErr:
    Dim strErrorID As String
    strErrorID = SaveError(Err.Number, "modFiles.WriteFile", Err.Description)
    On Error Resume Next
    Close lFile
    RaiseError strErrorID
End Sub

Public Function FileExists(strFilePath As String) As Boolean
On Error GoTo CatchErr
Dim fso As FileSystemObject

   Set fso = New FileSystemObject
   FileExists = fso.FileExists(strFilePath)
    Set fso = Nothing

Exit Function
CatchErr:
    Err.Raise Err.Number, "modFiles.FileExists", Err.Description
End Function
