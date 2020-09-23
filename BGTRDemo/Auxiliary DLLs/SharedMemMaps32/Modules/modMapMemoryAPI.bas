Attribute VB_Name = "modMapMemoryAPI"
'********************************************************************************************
' modMapMemoryAPI Module Definition
'
'  This is a standard module for the API declarations
'   needed to use shared/mapped memory.
'
'********************************************************************************************
Option Explicit

'Miscellaneous API constants
Public Const FILE_MAP_WRITE = &H2
Public Const PAGE_READWRITE = 4&
Public Const ERROR_ALREADY_EXISTS = 183&
Public Const WAIT_OBJECT_0 = 0&

'Function declarations
Public Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" _
                                       (ByVal hFile As Long, _
                                        ByVal lpFileMappingAttributes As Long, _
                                        ByVal flProtect As Long, _
                                        ByVal dwMaximumSizeHigh As Long, _
                                        ByVal dwMaximumSizeLow As Long, _
                                        ByVal lpName As String) _
                                    As Long

Public Declare Function MapViewOfFile Lib "kernel32" _
                                       (ByVal hFileMappingObject As Long, _
                                        ByVal dwDesiredAccess As Long, _
                                        ByVal dwFileOffsetHigh As Long, _
                                        ByVal dwFileOffsetLow As Long, _
                                        ByVal dwNumberOfBytesToMap As Long) _
                                    As Long

Public Declare Function UnmapViewOfFile Lib "kernel32" _
                                         (lpBaseAddress As Any) _
                                    As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                               (Destination As Any, _
                                Source As Any, _
                                ByVal Length As Long)

Public Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" _
                                       (ByVal lpMutexAttributes As Long, _
                                        ByVal bInitialOwner As Long, _
                                        ByVal lpName As String) _
                                    As Long

Public Declare Function CreateSemaphore Lib "kernel32" Alias "CreateSemaphoreA" _
                                       (ByVal lpSemaphoreAttributes As Long, _
                                        ByVal lInitialCount As Long, _
                                        ByVal lMaximumCount As Long, _
                                        ByVal lpName As String) _
                                    As Long

Public Declare Function ReleaseMutex Lib "kernel32" _
                                        (ByVal hMutex As Long) _
                                    As Long

Public Declare Function ReleaseSemaphore Lib "kernel32" _
                                       (ByVal hSemaphore As Long, _
                                        ByVal lReleaseCount As Long, _
                                        lpPreviousCount As Long) _
                                    As Long

Public Declare Function OpenSemaphore Lib "kernel32" Alias "OpenSemaphoreA" _
                                       (ByVal dwDesiredAccess As Long, _
                                        ByVal bInheritHandle As Long, _
                                        ByVal lpName As String) _
                                    As Long

Public Declare Function WaitForSingleObject Lib "kernel32" _
                                       (ByVal hHandle As Long, _
                                        ByVal dwMilliseconds As Long) _
                                    As Long

Public Declare Function CloseHandle Lib "kernel32" _
                                        (ByVal hObject As Long) _
                                    As Long

Public Declare Function OpenMutex Lib "kernel32" Alias "OpenMutexA" _
                                        (ByVal dwDesiredAccess As Long, _
                                         ByVal bInheritHandle As Long, _
                                         ByVal lpName As String) _
                                    As Long


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Table from the MSDN

'If the argument is                                                             Declare it as
'-----------------------------------------------------------------------------------------------------------------------------
'standard C string (LPSTR, char far *)                                 ByVal S$
'Visual Basic string (see note)                                            S$
'integer (WORD, HANDLE, int)                                           ByVal I%
'pointer to an integer (LPINT, int far *)                                  I%
'long (DWORD, unsigned long)                                           ByVal L&
'pointer to a long (LPDWORD, LPLONG, DWORD far *)        L&
'standard C array (A[ ])                                                      base type array (no ByVal)
'Visual Basic array (see note)                                             A()
'struct (typedef)                                                                  S As Struct


