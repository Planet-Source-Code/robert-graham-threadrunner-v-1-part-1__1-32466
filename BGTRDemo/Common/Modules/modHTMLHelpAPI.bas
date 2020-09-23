Attribute VB_Name = "modHTMLHelpAPI"
'********************************************************************************************
'  modHTMLHelpAPI Module Definition
'
'   This is a standard module for the HTML Help API declares
'
'********************************************************************************************
Option Explicit

Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_HELP_CONTEXT = &HF


Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" _
(ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, _
dwData As Any) As Long


