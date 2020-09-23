Attribute VB_Name = "modTaskDispatcher"
'********************************************************************************************
'   modTaskDispatcher Standard Module Definition
'
'   This module defines constants used for Win32 API declares
'   to find hidden window and the function called by task result recipients
'   to raise task errors
'
'********************************************************************************************
Option Explicit

'Public constants for location of hidden window
Public Const MAIN_MDI_CLASS As String = "ThunderRT6MDIForm"
Public Const MAIN_PICBOX_CLASS As String = "ThunderRT6PictureBoxDC"
Public Const MAIN_TEXTBOX_CLASS As String = "ThunderRT6TextBox"

'Stub out the above and use these instead in order to run in the IDE
'Public Const MAIN_MDI_CLASS As String = "ThunderMDIForm"
'Public Const MAIN_PICBOX_CLASS As String = "ThunderPictureBoxDC"
'Public Const MAIN_TEXTBOX_CLASS As String = "ThunderTextBox"


Public Const MAIN_PICBOX_NAME As String = ""

'This function returns the hwnd of the hidden window
Public Function FindThreadTextWindow() As Long
On Error GoTo CatchErr
Dim hwnd_Main As Long
Dim hwnd_PicBox As Long
Dim hwnd_MainText As Long

    hwnd_Main = FindWindow(MAIN_MDI_CLASS, frmMDIMain.Caption)
    
    hwnd_PicBox = FindWindowEx(hwnd_Main, 0&, MAIN_PICBOX_CLASS, vbNullChar)
    
    FindThreadTextWindow = FindWindowEx(hwnd_PicBox, 0, MAIN_TEXTBOX_CLASS, vbNullChar)
    If FindThreadTextWindow = 0 Then
        'Raise an error!
    End If
    
Exit Function
CatchErr:

End Function

'Public function to allow the Dispatcher and Recipients to decipher an error
'memento returned from the ThreadRunner
Public Sub RaiseTaskError(ByRef arTaskMemento() As Byte)

Dim pbTemp As PropertyBag
Dim strErrorInfo As String
Dim tmpErrString As ErrorString_type
Dim tmpError As Error_type

    Set pbTemp = New PropertyBag
    With pbTemp
        .Contents = arTaskMemento
        strErrorInfo = .ReadProperty("ErrorInfo")
    End With
    
    Set pbTemp = Nothing
    
    tmpErrString.ErrorString = strErrorInfo
    LSet tmpError = tmpErrString
    'Raise the error in calling function
    Err.Raise tmpError.Number, Trim$(tmpError.Source), Trim$(tmpError.Description)
End Sub
