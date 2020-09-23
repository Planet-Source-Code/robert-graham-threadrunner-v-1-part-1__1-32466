Attribute VB_Name = "modMessageAPI"
'********************************************************************************************
'   modMessageAPI Standard Module Definition
'
'   This module defines Win32 API declares used by the thread
'   controller class to post messages to a hidden client window to notify
'   the client when tasks are completed.
'
'********************************************************************************************
Option Explicit

Public Const WM_SETTEXT = &HC

Public Declare Function SendMessage Lib "user32" _
                          Alias "SendMessageA" (ByVal hwnd As Long, _
                                                              ByVal wMsg As Long, _
                                                              ByVal wParam As Long, _
                                                              ByRef lParam As Any) _
                                    As Long

