Attribute VB_Name = "modTimerAPI"
'********************************************************************************************
' modTimerAPI Module Definition
'
' This is a standard module for the API declarations
' needed to create Win32 fire once timers.
'
'********************************************************************************************
Option Explicit

Public Declare Function SetTimer Lib "user32" _
                                        (ByVal hwnd As Long, _
                                         ByVal nIDEvent As Long, _
                                         ByVal uElapse As Long, _
                                         ByVal lpTimerFunc As Long) _
                                    As Long

Public Declare Function KillTimer Lib "user32" _
                                       (ByVal hwnd As Long, _
                                        ByVal nIDEvent As Long) _
                                   As Long
 

 


