Attribute VB_Name = "modMemoryAPI"
'********************************************************************************************
'   modMemoryAPI Standard Module Definition
'
'   This module defines Win32 API declares used by the thread
'   controller class to control referencing and modify memory
'   addresses.
'
'********************************************************************************************
Option Explicit

Public Declare Function InterlockedIncrement Lib "kernel32" _
                                    (ByVal lpAddend As Long) _
                      As Long


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                                (Destination As Any, _
                                Source As Any, _
                                ByVal Length As Long)

