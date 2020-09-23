VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBGTRMainFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************************************
'   CBGTRMainFactory Class Definition
'
'   This class defines a concrete Main Factory which implements the abstract
'   IBGTRMainFactory interface.   The class creates concrete classes to
'   implement the IBGTRWorkFactory interface, based upon the parameters
'   supplied by the client code.
'
'   Instancing is set to:  5 - MultiUse
'********************************************************************************************
Option Explicit

Implements IBGTRMainFactory

'The input parameter is specified as an array to support as much
' information as may be required to identify the correct DLL to
'provide the secondary factory.

'The concrete factory is generally specific to an application or group
' of applications.  This factory is specific to the demo app.
'Only a single dll will be used.  Some common possibilities
'are included for illustration, but have been stubbed out.

Private Function IBGTRMainFactory_CreateWorkFactory(arFactoryDescriptors() As String) As IBGTRWorkFactory

On Error GoTo CatchErr
Dim strFactoryDescriptor As String

    'Only need one for this application
    strFactoryDescriptor = arFactoryDescriptors(0)
    
    Select Case strFactoryDescriptor
                
        Case "DBTasks"
            'Create the concrete WorkFactory instance
            
            '************UnComment this line during build process**********
            'Set IBGTRMainFactory_CreateWorkFactory = New BGTRDBWorks.CBGTRDBWorkFactory
            
        Case "FileTasks"
            'Could implement file search, transfer etc
            'example:  Set IBGTRMainFactory_CreateWorkFactory = YourFileTaskDLL.FileWorkFactory

        Case "StringTasks"
            'Could implement time consuming string operations or text search operations
            'example:  Set IBGTRMainFactory_CreateWorkFactory = YourStringTaskDLL.StringWorkFactory

        Case "PrintTasks"
            'Good BG candidate in principle, logistics may not be viable
            'example:  Set IBGTRMainFactory_CreateWorkFactory = YourPrintTaskDLL.PrintWorkFactory

        Case "AnySundryLongRunningTask"
            'etc. etc. etc. etc.

        Case Else
            'Raise an Error
            Err.Raise 5             'Invalid procedure call
    End Select
    
Exit Function
CatchErr:
'No cleanup required if we errored - just return the error to the Worker
    Err.Raise Err.Number, "CBGTRMainFactory.CreateWorkFactory", Err.Description
End Function
