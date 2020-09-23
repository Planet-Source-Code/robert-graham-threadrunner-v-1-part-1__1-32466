Attribute VB_Name = "modMain"
Option Explicit


'Sub Main is executed when the DLL loads and simply initializes
'the default no data byte array one time for the life of the DLL

Sub Main()
    ReDim g_arNoData(0)
    g_arNoData(0) = BYTE_OFF
End Sub


