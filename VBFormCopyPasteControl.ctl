VERSION 5.00
Begin VB.UserControl VBFormCopyPasteControl 
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   360
   InvisibleAtRuntime=   -1  'True
   Picture         =   "VBFormCopyPasteControl.ctx":0000
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   24
   ToolboxBitmap   =   "VBFormCopyPasteControl.ctx":00E2
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Begin VB.Menu mnuFormCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuFormPaste 
         Caption         =   "&Paste"
      End
   End
End
Attribute VB_Name = "VBFormCopyPasteControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_PasteEnabled = 0
Const m_def_PasteMenuText = "&Paste"
Const m_def_CopyMenuText = "&Copy"
Const m_def_ParentFormName = "MyFormName"
'Property Variables:
Dim m_PasteEnabled As Boolean
Dim m_PasteMenuText As String
Dim m_CopyMenuText As String
Dim m_ParentFormName As String


'\\ My clipboard format....
Private mCFCustom As Long
'\\ Subclassing previous window proc
Private mOldWndProc As Long
'\\ Clipboard chain - previous viewer
Private mPrevCBChainWindow As Long

'--[Subclassing stuff]--------------------------------------
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = (-4)
'------------------------------------------------------------

Public Event ClipboardChanged()

Friend Sub ClipboardChangedEvent()

RaiseEvent ClipboardChanged

End Sub


Friend Property Get ClipboardFormat() As Long

    If mCFCustom = 0 Then
        mCFCustom = RegisterClipboardFormat(m_ParentFormName)
    End If
    ClipboardFormat = mCFCustom
    
End Property


Private Sub CopyForm()



End Sub

Public Property Get IsFormDataInClipboard() As Boolean

    IsFormDataInClipboard = IsClipboardFormatAvailable(ClipboardFormat)
    
End Property


Public Property Get Key() As String
    
    Key = Makekey(UserControl.hwnd)
    
End Property




Friend Property Get OldWndProc() As Long
    OldWndProc = mOldWndProc
End Property

Public Property Get PasteEnabled() As Boolean
    If Ambient.UserMode Then Err.Raise 393
    PasteEnabled = m_PasteEnabled
End Property

Public Property Let PasteEnabled(ByVal New_PasteEnabled As Boolean)
    If Ambient.UserMode Then Err.Raise 393
    m_PasteEnabled = New_PasteEnabled
    PropertyChanged "PasteEnabled"
End Property

Private Sub PasteForm()


End Sub

Public Property Get PasteMenuText() As String
    PasteMenuText = m_PasteMenuText
End Property

Public Property Let PasteMenuText(ByVal New_PasteMenuText As String)
    m_PasteMenuText = New_PasteMenuText
    mnuFormPaste.Caption = m_PasteMenuText
    
    PropertyChanged "PasteMenuText"
End Property

Public Property Get CopyMenuText() As String
    CopyMenuText = m_CopyMenuText
End Property

Public Property Let CopyMenuText(ByVal New_CopyMenuText As String)
    m_CopyMenuText = New_CopyMenuText
    mnuFormCopy.Caption = m_CopyMenuText
    
    PropertyChanged "CopyMenuText"
End Property

Public Property Get ParentFormName() As String
    ParentFormName = m_ParentFormName
End Property

Public Property Let ParentFormName(ByVal New_ParentFormName As String)
    m_ParentFormName = New_ParentFormName
    PropertyChanged "ParentFormName"
End Property

Friend Property Let PreviousCBChainWindow(ByVal wndlast As Long)
    mPrevCBChainWindow = wndlast
End Property

Friend Property Get PreviousCBChainWindow() As Long
    PreviousCBChainWindow = mPrevCBChainWindow
End Property


'\\ --[ShowMenu]------------------------------------------
'\\ Called by the subclass thing to get this control to
'\\ show the copy/paste menu
'\\ ------------------------------------------------------
Friend Sub ShowMenu()

mnuFormPaste.Enabled = (IsFormDataInClipboard And PasteEnabled)

PopupMenu mnuPopup

End Sub

Public Property Get WindowHandle() As Long
    WindowHandle = UserControl.hwnd
End Property

Friend Property Let WindowProc(ByVal lpfnNew As Long)

If lpfnNew = 0 Then
    If mOldWndProc <> 0 Then
        Call SetWindowLong(UserControl.hwnd, GWL_WNDPROC, mOldWndProc)
    End If
Else
    mOldWndProc = SetWindowLong(UserControl.hwnd, GWL_WNDPROC, lpfnNew)
End If

End Property

Private Sub mnuFormCopy_Click()

Call CopyForm

End Sub

Private Sub mnuFormPaste_Click()

Call PasteForm

End Sub


Private Sub UserControl_Initialize()

    Call Startup
    
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_PasteEnabled = m_def_PasteEnabled
    
    PasteMenuText = m_def_PasteMenuText
    CopyMenuText = m_def_CopyMenuText
    
    m_ParentFormName = m_def_ParentFormName
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_PasteEnabled = PropBag.ReadProperty("PasteEnabled", m_def_PasteEnabled)
    
    PasteMenuText = PropBag.ReadProperty("PasteMenuText", m_def_PasteMenuText)
    CopyMenuText = PropBag.ReadProperty("CopyMenuText", m_def_CopyMenuText)
    
    m_ParentFormName = PropBag.ReadProperty("ParentFormName", m_def_ParentFormName)
    '\\ If we are in runtime mode, register the clipboard format...
    If Not (Ambient.UserMode) Then
        mCFCustom = RegisterClipboardFormat(m_ParentFormName)
    End If

    If Ambient.UserMode Then
        Call SubclassControl(Me)
    End If
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("PasteEnabled", m_PasteEnabled, m_def_PasteEnabled)
    Call PropBag.WriteProperty("PasteMenuText", m_PasteMenuText, m_def_PasteMenuText)
    Call PropBag.WriteProperty("CopyMenuText", m_CopyMenuText, m_def_CopyMenuText)
    Call PropBag.WriteProperty("ParentFormName", m_ParentFormName, m_def_ParentFormName)
End Sub

