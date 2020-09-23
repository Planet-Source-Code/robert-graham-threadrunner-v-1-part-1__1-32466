Attribute VB_Name = "modWindowsAPI"
'********************************************************************************************
'   modWindowsAPI Standard Module Definition
'
'   This module defines Win32 API declares used to find hidden window
'   and for auxiliary support such as managing MDI forms
'
'********************************************************************************************
Option Explicit

Public Enum WindowPos
  vbTopMost = -1&
  vbNotTopMost = -2&
  vbTop = 0&
  vbBottom = 1&
End Enum

    
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type



Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal _
hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                           lpRect As RECT) As Long

Public Declare Function FlashWindow Lib "user32" _
                       (ByVal hwnd As Long, _
                       ByVal bInvert As Long) As Long

Public Declare Function GetCaretBlinkTime Lib "user32" () As Long


