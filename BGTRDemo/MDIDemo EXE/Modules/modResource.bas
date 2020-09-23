Attribute VB_Name = "modResource"
'********************************************************************************************
'  modResource Module Definition
'
'   This is a standard module for resource file ID declarations
'
'********************************************************************************************
Option Explicit

'Declare constants for accessing the memory spaces

'Number of data map files (= max number of concurrent tasks/workers)
Public Const RES_MAP_NUM As Long = 3998
'Name will be suffixed by hInstance + map file number
Public Const RES_MEM_NAME As Long = 3999
'Map sizes are declared starting at this ID
Public Const RES_MEM_START As Long = 4000

'Declare constant for resource ID of the filemove.avi resource
Public Const RES_FILEMOVE_AVI As Long = 3000

'Constant for demo icon
Public Const RES_DEMO_ICON As Long = 2999

