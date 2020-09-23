VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJetProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CJetProc Class Definition
'   Class defines methods for creating and manipulating stored procedures
'   in Jet (Access) databases using ADOX methods
'
'   Requires a project reference to the MSADOX Dll
'
'   Instancing is set to:  5 - MultiUse
''
'********************************************************************************************
Option Explicit

Public Enum JetProcParameter_enum
    jppcol_ParameterName = 0
    jppcol_ParameterType = 1
End Enum


Public Sub DropAllJetProcedures(cnConnection As Connection)
On Error GoTo CatchErr
Dim cat As ADOX.Catalog
Dim SpProc As ADOX.Procedure
Dim ICount As Long

    Set cat = New ADOX.Catalog
    ' Open the Catalog
    Set cat.ActiveConnection = cnConnection
        'The Catalog.Procedures collection does NOT support For Each ..Next enumeration!
        For ICount = 0 To cat.Procedures.Count - 1
            With cat.Procedures
                Set SpProc = .Item(0)
                .Delete SpProc.Name
            End With
        Set SpProc = Nothing
        Next ICount
        
  Set cat.ActiveConnection = Nothing
  Set cat = Nothing
Exit Sub
CatchErr:
      Err.Raise Err.Number, Err.Source & " in VBADOTools.CJetProc.DropAllJetProcedures", Err.Description
End Sub

Public Sub DropJetProcedure(cnConnection As Connection, ByRef strProcName As String)
On Error GoTo CatchErr
Dim cat As ADOX.Catalog
Dim SpProc As ADOX.Procedure
Dim lcount As Long

    Set cat = New ADOX.Catalog
    ' Open the Catalog
    Set cat.ActiveConnection = cnConnection
        'The Catalog.Procedures collection does NOT support For Each ..Next enumeration!
        For lcount = 0 To cat.Procedures.Count - 1
            With cat.Procedures
                Set SpProc = .Item(lcount)
                If SpProc.Name = strProcName Then
                    .Delete SpProc.Name
                    Set SpProc = Nothing
                    Exit For
                End If
            End With
        Set SpProc = Nothing
        Next lcount
        
  Set cat.ActiveConnection = Nothing
  Set cat = Nothing
      
Exit Sub
CatchErr:
      Err.Raise Err.Number, Err.Source & " in VBADOTools.CJetProc.DropJetProcedure", Err.Description
End Sub


Public Sub CreateJetProcedure(cnConnection As Connection, strProcName As String, strCommand As String, lngParameterCount As Long, Optional arParameters As Variant)
On Error GoTo CatchErr
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim cat As ADOX.Catalog
Dim strParameters As String
Dim i As Long


    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cnConnection
    
    ' Create the parameterized command (Microsoft Jet specific)
    
    strParameters = vbNullString
    If lngParameterCount > 0 Then
        strParameters = "PARAMETERS "
        For i = 1 To lngParameterCount - 1
            strParameters = strParameters & " " & arParameters(i - 1, jppcol_ParameterName) & " " & arParameters(i - 1, jppcol_ParameterType) & ","
        Next i
        'i = lngParameterCount
        strParameters = strParameters & " " & arParameters(i - 1, jppcol_ParameterName) & " " & arParameters(i - 1, jppcol_ParameterType) & ";"
    End If
    
    cmd.CommandText = strParameters & strCommand
    
    Set cat = New ADOX.Catalog
    ' Open the Catalog
    Set cat.ActiveConnection = cnConnection

    ' Create the new Procedure
    cat.Procedures.Append strProcName, cmd
    Set cat.ActiveConnection = Nothing
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    
Exit Sub
CatchErr:
      Err.Raise Err.Number, Err.Source & " in VBADOTools.CJetProc.CreateJetProcedure", Err.Description
End Sub


Public Function EnumerateJetProcedures(cnConnection As Connection) As Collection
On Error GoTo CatchErr
Dim cmd As ADODB.Command
Dim cat As ADOX.Catalog
Dim SpProc As ADOX.Procedure
Dim strProcedure As String
Dim colProcs As Collection
Dim lcount As Long

    Set cat.ActiveConnection = cnConnection
    Set colProcs = New Collection
    
    With cat.Procedures
        For lcount = 0 To .Count - 1
            Set SpProc = .Item(lcount)
            Set cmd = SpProc.Command
            strProcedure = cmd.Name & " : " & cmd.CommandText
            colProcs.Add strProcedure
            Set cmd = Nothing
            Set SpProc = Nothing
        Next lcount
    End With
    
    Set cat.ActiveConnection = Nothing
    Set EnumerateJetProcedures = colProcs
    Set colProcs = Nothing
    
Exit Function
CatchErr:
      Err.Raise Err.Number, Err.Source & " in VBADOTools.CJetProc.EnumerateJetProcedures", Err.Description
End Function


'The following are valid Jet types for declaring Jet Stored procedures.
'Comments:  Use TEXT for strings with no length specification.  Use TEXT(n) for an equivalent to the CHAR type.
'                   Use INTEGER for longs.  Use BIT for Booleans

'BINARY 1 byte per character Any type of data may be stored in a field of this type. No translation of the data (for example, to text) is made. How the data is input in a binary field dictates how it will appear as output.
'BIT 1 byte Yes and No values and fields that contain only one of two values.
'TINYINT 1 byte An integer value between 0 and 255.
'MONEY 8 bytes A scaled integer between  – 922,337,203,685,477.5808 and 922,337,203,685,477.5807.
'DATETIME 8 bytes A date or time value between the years 100 and 9999.
'UNIQUEIDENTIFIER 128 bits A unique identification number used with remote procedure calls.
'REAL 4 bytes A single-precision floating-point value with a range of  – 3.402823E38 to  – 1.401298E-45 for negative values, 1.401298E-45 to 3.402823E38 for positive values, and 0.
'FLOAT 8 bytes A double-precision floating-point value with a range of  – 1.79769313486232E308 to  – 4.94065645841247E-324 for negative values, 4.94065645841247E-324 to 1.79769313486232E308 for positive values, and 0.
'SMALLINT 2 bytes A short integer between  – 32,768 and 32,767. (See Notes)
'INTEGER 4 bytes A long integer between  – 2,147,483,648 and 2,147,483,647. (See Notes)
'DECIMAL 17 bytes An exact numeric data type that holds values from 1028 - 1 through - 1028 - 1. You can define both precision (1 - 28) and scale (0 - defined precision). The default precision and scale are 18 and 0, respectively.
'TEXT 2 bytes per character Zero to a maximum of 2.14 gigabytes.
'IMAGE As required Zero to a maximum of 2.14 gigabytes. Used for OLE objects.
'CHARACTER 2 bytes per character Zero to 255 characters.




