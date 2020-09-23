VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMemoryMapFileTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CMemoryMapFileTest Class Definition
'
'   This class encapsulates the entire process of creating shared memory,
'   creating synchronization objects, and reading and writing to the memory.
'
'   A class instance represents a single shared memory file mapping.
'
'   This is a special test version of the class in SharedMemMaps32 DLL
'********************************************************************************************
Option Explicit

'Public enum of error codes
Public Enum SharedMemMaps32ErrorCodes_enum
    sharedmemerr_CreateFile = 50000
    sharedmemerr_CreateMap = 50001
    sharedmemerr_CreateMutex = 50002
    sharedmemerr_CreateSemaphore = 50003
End Enum


'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^  Magic numbers for file mapping API calls
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Declare a constant of 0 for the security attributes
Private Const SEC_ATTRIBUTES As Long = 0

'Declare a constant of -1 for the hFile handle because these functions
'will read/write blocks of memory directly, with no use of a file.
Private Const HFILE_NONE As Long = -1
'Declare a constant of 0 for the memory File maximum size highword
Private Const MAP_SIZE_HIGH As Long = 0
'Declare constants for memory offsets - none are used
Private Const OFFSET_HIGH As Long = 0
Private Const OFFSET_LOW As Long = 0
''Declare constants for memory access - all of it
Private Const ACCESS_ALL As Long = 0

'Declare const for mutex initial ownership
Private Const OWN_MUTEX_NO As Boolean = False
'Declare const for maximum wait time for a mutex - msecs
Private Const WAIT_MAX_MUTEX As Long = 5000

'Declare const for semaphore access priviliges - full access
Private Const SEM_ACCESS_FULL As Long = -1
'Declare const for semaphore inheritence - cannot be inherited
Private Const CANT_INHERIT As Boolean = False
'Declare const for maximum wait time for a semaphore - msecs
Private Const WAIT_MAX_SEM As Long = 1000

'Declare const for sleep time if required - msecs
Private Const NAP_TIME As Long = 100

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Declares for various objects and handles
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Memory File handles
Private m_hwndMemoryMap As Long
Private m_pMemoryMap As Long
'
'Synchronization object handles
Private m_hwndMutex As Long
Private m_hwndSemaphore As Long

'Maximum number of semaphores to create
'also represents the maximum threads that can
'use a given shared memory File.
Private All_Semaphores As Long

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Memory map file creation function
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Public Sub MapMemory(strShareName As String, lngMaxLength As Long)
On Error GoTo CatchErr
Dim strErrorSource As String
Dim lngMapError As Long
Dim strName As String
    
    'Declare error source
    strErrorSource = "SharedMemMaps32.CMemoryMapFile.MapMemory"
    'Ensure the Err object is clean when we start
    Err.Clear
    
    strName = strShareName & ".map"
        
    'lngMaxLength is entered as "KB" so calculate correct memory size in multiples of 1024
    lngMaxLength = lngMaxLength * BYTES_KB

    'Attempt to create the mapped memory File
    m_hwndMemoryMap = CreateFileMapping(HFILE_NONE, _
                                                                  SEC_ATTRIBUTES, _
                                                                  PAGE_READWRITE, _
                                                                  MAP_SIZE_HIGH, _
                                                                  lngMaxLength, _
                                                                  strName$)

    If m_hwndMemoryMap = 0 Then
        'Raise an error
       Err.Raise vbObjectError + sharedmemerr_CreateFile, strErrorSource, LoadResString(sharedmemerr_CreateFile)
    End If
    
    'Cache the last error info
    lngMapError = Err.LastDllError

    m_pMemoryMap = MapViewOfFile(m_hwndMemoryMap, _
                                                       FILE_MAP_WRITE, _
                                                       OFFSET_HIGH, _
                                                       OFFSET_LOW, _
                                                       ACCESS_ALL)
    
    If m_pMemoryMap = 0 Then
         'Raise an error
       Err.Raise vbObjectError + sharedmemerr_CreateMap, strErrorSource, LoadResString(sharedmemerr_CreateMap)
    End If
    
    'Initialize the memory with no data
    If Not lngMapError = ERROR_ALREADY_EXISTS Then
        CopyMemory ByVal m_pMemoryMap, NO_DATA, BYTES_LONG
    End If

    'Try to create the mutex.
    m_hwndMutex = CreateMutex(SEC_ATTRIBUTES, _
                                                OWN_MUTEX_NO, _
                                                strShareName & ".mtx")

    'If the mutex could not be created, destroy all variables and bail.
    If m_hwndMutex = 0 Then
         'Raise an error
       Err.Raise vbObjectError + sharedmemerr_CreateMutex, strErrorSource, LoadResString(sharedmemerr_CreateMutex)
    End If

    'Attempt to open a semaphore.

m_hwndSemaphore = OpenSemaphore(SEM_ACCESS_FULL, _
                                                         CANT_INHERIT, _
                                                         (strShareName & ".sem"))

'If the semaphore could not be opened, create one.
If m_hwndSemaphore = 0 Then
    m_hwndSemaphore = CreateSemaphore(SEC_ATTRIBUTES, _
                                                               All_Semaphores, _
                                                               All_Semaphores, _
                                                               strShareName & ".sem")

    'If the semaphore could not be created, destroy all variables and bail.
    If m_hwndSemaphore = 0 Then
         'Raise an error
       Err.Raise vbObjectError + sharedmemerr_CreateSemaphore, strErrorSource, LoadResString(sharedmemerr_CreateSemaphore)
    End If
End If

Exit Sub

CatchErr:
Dim strErrorID As String
    'First save the error info
    strErrorID = SaveError(Err.Number, Err.Source & " Raised from " & strErrorSource, Err.Description)
    'Now clean up!
    On Error Resume Next
    'Close any and all handles
    UnmapViewOfFile m_pMemoryMap
    CloseHandle m_hwndMemoryMap
    CloseHandle m_hwndMutex
    CloseHandle m_hwndSemaphore
    'Raise the saved error
    RaiseError strErrorID
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Functions to access and use memory map files
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public Sub WriteMemory(Data() As Byte)
On Error GoTo CatchErr
Dim arContents() As Byte
Dim lngLength As Long
Dim lngRetVal As Long
Dim lngSemaphoreCount As Long
    
    'Set array to parameter Data array
    arContents = Data
    
    'Determine the size of the array
    lngLength = UBound(arContents) + 1
    
    'Grab the mutex.
    Do
        lngRetVal = WaitForSingleObject(m_hwndMutex, WAIT_MAX_MUTEX)
    Loop Until lngRetVal = WAIT_OBJECT_0
    
    'Wait until semaphore count signals no readers are left in mapped memory.
    Do
        Do
        Loop Until WaitForSingleObject(m_hwndSemaphore, WAIT_MAX_SEM) = WAIT_OBJECT_0
            ReleaseSemaphore m_hwndSemaphore, 1, lngSemaphoreCount
    Loop Until lngSemaphoreCount = All_Semaphores - 1
    
    'Copy the size of the array into first 4 bytes of memory.
    CopyMemory ByVal m_pMemoryMap, lngLength, BYTES_LONG
    
    'Copy the array starting at 5th byte of memory.
    CopyMemory ByVal (m_pMemoryMap + BYTES_LONG), arContents(0), lngLength
    
    'Release the mutex.  Readers are now free to enter memory.
    ReleaseMutex m_hwndMutex
Exit Sub
CatchErr:
'Not much we can do.  We could release the mutex to be safe, but
'we don't know if we got a semaphore - releasing could screw the count.
'We're basically screwed - do nothing, just propagate the error back
    Err.Raise Err.Number, "SharedMemMaps32.CMemoryMapFile.WriteMemory"
End Sub

Public Function ReadMemory() As Byte()
On Error GoTo CatchErr
Dim lngLength As Long
Dim arContents() As Byte
Dim lngSemaphoreCount As Long
Dim lngRetVal As Long
    
    'Get the mutex, to make sure writer does not have it.
     Do
        lngRetVal = WaitForSingleObject(m_hwndMutex, WAIT_MAX_MUTEX)
        If lngRetVal <> WAIT_OBJECT_0 Then
            Sleep NAP_TIME      'Take a nap, and try again
        End If
    Loop Until lngRetVal = WAIT_OBJECT_0
    
    'Grab a semaphore while still holding the mutex.
    
    Do
    Loop Until WaitForSingleObject(m_hwndSemaphore, WAIT_MAX_SEM) = WAIT_OBJECT_0
    
    'Now we are in, and semaphored.  Release the mutex immediately, so next reader
    'can grab it and get in.
    
    ReleaseMutex m_hwndMutex
    
    'Time to get to work!
    
    'Get the first 4 bytes.
    
    CopyMemory lngLength, ByVal m_pMemoryMap, BYTES_LONG
    
    'Check the value
    
    If lngLength = 0 Then
        'Nothing here! Let's initialize the function return and split!
        ReDim arContents(NO_DATA)
        'Release the semaphore
        ReleaseSemaphore m_hwndSemaphore, 1, lngSemaphoreCount
        Exit Function
    End If
    
    'Read the memory.
    ReDim arContents(lngLength - 1)
    CopyMemory arContents(0), ByVal (m_pMemoryMap + BYTES_LONG), lngLength
    
    'Time to split!
    'Release the semaphore
    ReleaseSemaphore m_hwndSemaphore, 1, lngSemaphoreCount
    'All done!
    ReadMemory = arContents
Exit Function
CatchErr:
'Same deal as before:
'We're basically screwed - do nothing, just propagate the error back
    Err.Raise Err.Number, "SharedMemMaps32.CMemoryMapFile.ReadMemory"
End Function

Public Sub ResetMemory()
On Error GoTo CatchErr
Dim lngLength As Long
Dim lngRetVal As Long
Dim lngSemaphoreCount As Long
    
    'Grab the mutex.
    Do
        lngRetVal = WaitForSingleObject(m_hwndMutex, WAIT_MAX_MUTEX)
    Loop Until lngRetVal = WAIT_OBJECT_0
    
    'Wait until semaphore count signals no readers are left in mapped memory.
    Do
        Do
            Loop Until WaitForSingleObject(m_hwndSemaphore, WAIT_MAX_SEM) = WAIT_OBJECT_0
            ReleaseSemaphore m_hwndSemaphore, 1, lngSemaphoreCount
    Loop Until lngSemaphoreCount = All_Semaphores - 1
    
    'Copy a zero into first 4 bytes of memory.
    CopyMemory ByVal m_pMemoryMap, NO_DATA, BYTES_LONG
    
    'Release the mutex.
    ReleaseMutex m_hwndMutex
Exit Sub
CatchErr:
'Seen this movie too!
'We're basically screwed - do nothing, just propagate the error back
    Err.Raise Err.Number, "SharedMemMaps32.CMemoryMapFile.ResetMemory"
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Class constructor/destructor
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub Class_Initialize()
On Error GoTo CatchErr
    'Retrieve max semaphores from the res file
    All_Semaphores = CLng(LoadResString(RES_MAX_SEMAPHORES))
Exit Sub
CatchErr:
'Should not happen!
    Err.Raise Err.Number, "SharedMemMaps32.CMemoryMapFile.Class_Initialize"
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    UnmapViewOfFile m_pMemoryMap
    CloseHandle m_hwndMemoryMap
    CloseHandle m_hwndMutex
    CloseHandle m_hwndSemaphore
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'^ Special functions added only for test purposes!
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Public Sub HoldMutex()
Dim lngRetVal As Long
'Grab the mutex.
    Do
        lngRetVal = WaitForSingleObject(m_hwndMutex, WAIT_MAX_MUTEX)
    Loop Until lngRetVal = WAIT_OBJECT_0
End Sub

Public Sub DropMutex()
 'Release the mutex.
    ReleaseMutex m_hwndMutex
End Sub

Public Sub HoldSemaphore()
Dim lngRetVal As Long
    
    'Get the mutex, to make sure writer does not have it.
     Do
        lngRetVal = WaitForSingleObject(m_hwndMutex, WAIT_MAX_MUTEX)
        If lngRetVal <> WAIT_OBJECT_0 Then
            Sleep NAP_TIME      'Take a nap, and try again
        End If
    Loop Until lngRetVal = WAIT_OBJECT_0
    
    'Grab a semaphore while still holding the mutex.
    
    Do
    Loop Until WaitForSingleObject(m_hwndSemaphore, WAIT_MAX_SEM) = WAIT_OBJECT_0
    
    'Now we are in, and semaphored.  Release the mutex immediately, so next reader
    'can grab it and get in.
    
    ReleaseMutex m_hwndMutex
End Sub

Public Sub DropSemaphore()
Dim lngSemaphoreCount As Long
'Release the semaphore
 ReleaseSemaphore m_hwndSemaphore, 1, lngSemaphoreCount
End Sub

