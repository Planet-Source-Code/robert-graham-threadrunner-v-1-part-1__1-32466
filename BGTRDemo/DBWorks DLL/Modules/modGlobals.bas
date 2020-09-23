Attribute VB_Name = "modGlobals"
'********************************************************************************************
'   modGlobals Standard Module Definition
'
'   This module defines misc dll level scope globals
'
'********************************************************************************************
Option Explicit

'Constant to indicate frequency at which to check Cancel
Public Const CHECK_CANCEL_INCR As Long = 500

'Public default no data byte array
Public g_arNoData() As Byte

