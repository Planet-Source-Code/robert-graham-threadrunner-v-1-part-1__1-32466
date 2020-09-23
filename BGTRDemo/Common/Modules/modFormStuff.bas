Attribute VB_Name = "modFormStuff"
Option Explicit

Private lngVerticalOffset As Long

Public Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum

Public Const INVERT As Long = 1
Public Const NO_INVERT As Long = 0


Public Sub InitializeMDIClient()
On Error Resume Next
Dim ctrl As Control
    Set ctrl = frmMDIMain.pctMDIMain
    If Not ctrl Is Nothing Then
        lngVerticalOffset = ctrl.Height
        Set ctrl = Nothing
    Else
        lngVerticalOffset = 0
    End If
End Sub

Public Sub CloseMDIForm()
On Error Resume Next
Dim frm As Form
    Set frm = frmMDIMain.ActiveForm
    If Not frm Is Nothing Then
        Unload frm
        Set frm = Nothing
    End If
    
End Sub

Public Sub CloseAllMDIForms()
On Error Resume Next
Dim bIsMoreForms As Boolean
Dim frm As Form

    bIsMoreForms = True

    Do While bIsMoreForms
        Set frm = frmMDIMain.ActiveForm
        If frm Is Nothing Then
            bIsMoreForms = False
        Else
            bIsMoreForms = True
            Unload frm
        End If
    Loop
   
End Sub

Public Sub CenterForm(frm As Form)
Const FIFTHEEN As Long = 15
Const ONE_HALF As Double = 0.5
Dim x As Integer
Dim y As Integer
Dim ParentWdth As Integer
Dim ParentHgt As Integer
Dim bIsChildFrm As Boolean
Dim Rct As RECT

    On Error Resume Next
    
    If frm.MDIChild Then
        If Err = False Then
            On Error GoTo CatchErr
            bIsChildFrm = True
            Call GetClientRect(GetParent(frm.hWnd), Rct)
            ParentWdth = (Rct.Right - Rct.Left) * Screen.TwipsPerPixelY
            ParentHgt = (Rct.Bottom - Rct.Top) * Screen.TwipsPerPixelX
            x = (ParentWdth - frm.Width) * ONE_HALF
            y = (ParentHgt - frm.Height) * ONE_HALF
        End If
    End If
    
    If Not bIsChildFrm Then
        On Error GoTo CatchErr
        x = (Screen.Width - frm.Width) * ONE_HALF
        y = (Screen.Height - frm.Height) * ONE_HALF + lngVerticalOffset
    End If
    
    ' center the form and return True to indicate success
    frm.Move x, y
    Err = False
Exit Sub
    
CatchErr:
    MsgBox Err.Number & "  Centerform" & vbCrLf & Err.Description
    Err = False
End Sub


Public Function FormCount(ByVal frmName As String) As Long
    Dim frm As Form
    For Each frm In Forms
        If StrComp(frm.Name, frmName, vbTextCompare) = 0 Then
            FormCount = FormCount + 1
        End If
    Next
End Function

Public Sub HighlightControl(ctl As Control)
   ctl.SelStart = 0
   ctl.SelLength = Len(ctl)
End Sub

Public Sub SetFormPosition(frm As Form, Position As WindowPos)
   If Position = vbBottom Then
      SetWindowPos frm.hWnd, Position, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
   Else
      SetWindowPos frm.hWnd, Position, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
   End If
End Sub



