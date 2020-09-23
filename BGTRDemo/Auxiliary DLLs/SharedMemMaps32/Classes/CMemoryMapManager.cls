VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMemoryMapManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CMemoryMapManager Class Definition
'
'   This class instantiates MemoryMapFile objects
'   and manages the collection of MemoryMapFile objects.
'
'
'   Instancing is set to:  5 - MultiUse
'********************************************************************************************
Option Explicit

'Public enum of error codes
Public Enum SharedMemMaps32ErrorCodes_enum
    sharedmemerr_CreateFile = 50000
    sharedmemerr_CreateMap = 50001
    sharedmemerr_CreateMutex = 50002
    sharedmemerr_CreateSemaphore = 50003
End Enum


Public Enum FileMapArrayCol_enum
    mapcol_MapName = 0
    mapcol_MapKey = 1
    mapcol_MapSize = 2
End Enum


Private m_CMemoryMapFiles As CMemoryMapFiles


Public Sub MakeMapFiles(ByRef arMapFiles() As String)
On Error GoTo CatchErr
Dim i As Long
Dim Low As Long
Dim High As Long
Dim tmpFileName As String
Dim tmpFileKey As String
Dim tmpFileSize As Long

    Low = LBound(arMapFiles)
    High = UBound(arMapFiles)
    
    For i = Low To High
        
        tmpFileName = arMapFiles(i, mapcol_MapName)
        tmpFileKey = arMapFiles(i, mapcol_MapKey)
        tmpFileSize = CLng(arMapFiles(i, mapcol_MapSize))
        
        MakeMapFile tmpFileName, tmpFileKey, tmpFileSize
        
    Next i
Exit Sub
CatchErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub MakeMapFile(ByVal strShareName As String, ByVal sKey As String, ByVal lngLength As Long)
On Error GoTo CatchErr
Dim objCMemoryMapFile As CMemoryMapFile

    Set objCMemoryMapFile = New CMemoryMapFile
    objCMemoryMapFile.MapMemory strShareName, lngLength

    m_CMemoryMapFiles.Insert objCMemoryMapFile, sKey
    
Exit Sub
CatchErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub WriteMapMemory(ByVal sKey As String, Data() As Byte)
On Error GoTo CatchErr
Dim objCMemoryMapFile As CMemoryMapFile
    
    Set objCMemoryMapFile = m_CMemoryMapFiles.Item(sKey)
     'Write the memory by delegation to the class instance
     objCMemoryMapFile.WriteMemory Data()

Exit Sub
CatchErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ReadMapMemory(ByVal sKey As String) As Byte()
On Error GoTo CatchErr
Dim objCMemoryMapFile As CMemoryMapFile

    Set objCMemoryMapFile = m_CMemoryMapFiles.Item(sKey)
    'Read the memory by delegation to the class instance
    ReadMapMemory = objCMemoryMapFile.ReadMemory
    
Exit Function
CatchErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub CopyMapMemory(ByVal strTarget As String, ByVal strSource As String, Optional ByVal bZeroSource As Boolean = False)
On Error GoTo CatchErr
Dim objCMemoryMapFileTarget As CMemoryMapFile
Dim objCMemoryMapFileSource As CMemoryMapFile
    
    With m_CMemoryMapFiles
        Set objCMemoryMapFileTarget = .Item(strTarget)
        Set objCMemoryMapFileSource = .Item(strSource)
    End With
    
    'Read source, copy to target
    objCMemoryMapFileTarget.WriteMemory objCMemoryMapFileSource.ReadMemory
    If bZeroSource Then
        objCMemoryMapFileSource.ResetMemory
    End If
    
Exit Sub
CatchErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ZeroMapMemory(ByVal sKey As String)
On Error GoTo CatchErr
Dim objCMemoryMapFile As CMemoryMapFile

    Set objCMemoryMapFile = m_CMemoryMapFiles.Item(sKey)
     'Write the memory by delegation to the class instance
    objCMemoryMapFile.ResetMemory
    
Exit Sub
CatchErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Class constructor/destructor
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub Class_Initialize()
    Set m_CMemoryMapFiles = New CMemoryMapFiles
End Sub

Private Sub Class_Terminate()
Dim objCMemoryMapFile As CMemoryMapFile
On Error Resume Next

If Not m_CMemoryMapFiles Is Nothing Then
    For Each objCMemoryMapFile In m_CMemoryMapFiles
        Set objCMemoryMapFile = Nothing
    Next objCMemoryMapFile
    Set m_CMemoryMapFiles = Nothing
End If

End Sub
