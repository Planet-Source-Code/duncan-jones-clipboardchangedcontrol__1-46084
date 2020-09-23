Attribute VB_Name = "modCopyPasteFunctions"
Option Explicit

'--[Api Declarations]----------------------------------------
Public Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Public Declare Function GetClipboardData Lib "user32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As Long
Public Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Declare Function SetClipboardData Lib "user32" Alias "SetClipboardDataA" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
'------------------------------------------------------------

'--[Window messages]-----------------------------------------
Public Const WM_CLOSE = &H10
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_CHANGECBCHAIN = &H30D
'------------------------------------------------------------

'--[Subclassing API]-----------------------------------------
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'------------------------------------------------------------

'--[Clipboard viewer chain]----------------------------------
Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetClipboardViewer Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ChangeClipboardChain Lib "user32" (ByVal hwnd As Long, ByVal hWndNext As Long) As Long
'------------------------------------------------------------

Private mSubclassedControls As Collection

'--[LongFromLong]------------------------------------------
' To get past the silly "AddressOf" limitation
'---------------------------------------------------------
Private Function LongFromLong(ByVal lIn As Long) As Long
    LongFromLong = lIn
End Function


Public Sub SubclassControl(ByVal ctlIn As VBFormCopyPasteControl)

mSubclassedControls.Add ctlIn, Makekey(ctlIn.WindowHandle)
'\\ Set the windowproc and store the previous oie
ctlIn.PreviousCBChainWindow = SetClipboardViewer(ctlIn.WindowHandle)
ctlIn.WindowProc = LongFromLong(AddressOf VB_WndProc)

End Sub

Public Sub UnsubclassControl(ByVal ctlIn As VBFormCopyPasteControl)

'\\ Unset the windowproc to the previous one
ctlIn.WindowProc = 0
mSubclassedControls.Remove ctlIn.Key

End Sub
Public Function Makekey(ByVal hwnd As Long) As String

    Makekey = "FormCopyPasteControl:" & hwnd
    
End Function
Public Sub Startup()

If mSubclassedControls Is Nothing Then
    Set mSubclassedControls = New Collection
End If

End Sub
Public Function VB_WndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim ctlThis As VBFormCopyPasteControl

On Error Resume Next
Set ctlThis = mSubclassedControls.Item(Makekey(hwnd))

If ctlThis Is Nothing Then
    VB_WndProc = DefWindowProc(hwnd, wMsg, wParam, lParam)
Else
    Select Case wMsg
    Case WM_DRAWCLIPBOARD
        ctlThis.ClipboardChangedEvent
        '\\ pass the message on to the next clipboard viewer...
        If IsWindow(ctlThis.PreviousCBChainWindow) Then
            VB_WndProc = SendMessage(ctlThis.PreviousCBChainWindow, wMsg, wParam, lParam)
        End If
        
    Case WM_CHANGECBCHAIN
        '\\ wParam is the window being removed, lParam is the next in the chain...
        If ctlThis.PreviousCBChainWindow = wParam Then
            ctlThis.PreviousCBChainWindow = lParam
        End If
        If IsWindow(ctlThis.PreviousCBChainWindow) Then
            VB_WndProc = SendMessage(ctlThis.PreviousCBChainWindow, wMsg, wParam, lParam)
        End If
    
    Case WM_CLOSE
        If IsWindow(ctlThis.PreviousCBChainWindow) Then
            Call ChangeClipboardChain(ctlThis.WindowHandle, ctlThis.PreviousCBChainWindow)
        End If
        UnsubclassControl ctlThis
        VB_WndProc = DefWindowProc(hwnd, wMsg, wParam, lParam)
    Case Else
        VB_WndProc = CallWindowProc(ctlThis.OldWndProc, hwnd, wMsg, wParam, lParam)
    End Select
End If

End Function


