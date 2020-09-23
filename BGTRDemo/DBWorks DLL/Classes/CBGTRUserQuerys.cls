VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBGTRUserQuerys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CBGTRUserQuery"
Attribute VB_Ext_KEY = "Member0" ,"CBGTRUserQuery"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'   CBGTRUserQuerys Collection Class Definition
'   Class defines collection to contain UserQuery class instances
'   Standard VB collection implementation, with additional Insert method
'
'   Instancing is set to:  5 - MultiUse
''
'********************************************************************************************
Option Explicit

Private mCol As Collection

Public Sub Insert(ByRef oUserQuery As CBGTRUserQuery, Optional sKey As String = vbNullString)

    If Len(sKey) = 0 Then
        mCol.Add oUserQuery
    Else
        mCol.Add oUserQuery, sKey
    End If

End Sub

Public Function Add(ReturnsRecords As ReturnsRecords_enum, ParameterCount As Long, ADOParameters As Variant, SQLStatement As String, QueryType As QueryType_enum, Connection As CDBConnection, DBType As DBType_enum, Optional sKey As String) As CBGTRUserQuery
Dim objNewMember As CBGTRUserQuery
    
    Set objNewMember = New CBGTRUserQuery

    'set the properties passed into the method
    objNewMember.ReturnsRecords = ReturnsRecords
    objNewMember.ParameterCount = ParameterCount
    objNewMember.ADOParameters = ADOParameters
    objNewMember.SQLStatement = SQLStatement
    objNewMember.QueryType = QueryType
    objNewMember.DBType = DBType
   
    If IsObject(Connection) Then
        Set objNewMember.Connection = Connection
    Else
        objNewMember.Connection = Connection
    End If
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As CBGTRUserQuery
Attribute Item.VB_UserMemId = 0
  On Error Resume Next
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

