VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLinkTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************************
' CLinkTimers Collection Class Definition
'
' This class defines a collection which contains the CLinkTimer objects.
' Add is replaced by a new method - Insert - to allow CLinkTimer objects
' to be added directly. The objects are added using the TimerID as a key.
'
' Instancing is set to: 1 - Private
'********************************************************************************************
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Sub Insert(ByRef oCLinkTimer As CLinkTimer, Optional ByRef sKey As String)
    If Len(sKey) = 0 Then
        mCol.Add oCLinkTimer
    Else
        mCol.Add oCLinkTimer, sKey
    End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As CLinkTimer
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

