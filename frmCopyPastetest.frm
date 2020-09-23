VERSION 5.00
Object = "{921D6A27-9A6B-11D7-B3CB-00C04F84CB14}#2.0#0"; "VBFormCopyPaste.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VBFormCopyPaste.VBFormCopyPasteControl VBFormCopyPasteControl2 
      Left            =   3120
      Top             =   360
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   1815
   End
   Begin VBFormCopyPaste.VBFormCopyPasteControl VBFormCopyPasteControl1 
      Left            =   120
      Top             =   120
      _ExtentX        =   450
      _ExtentY        =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub VBFormCopyPasteControl1_ClipboardChanged()

Debug.Print "Changed 1 - " & VBFormCopyPasteControl1.IsFormDataInClipboard

End Sub

Private Sub VBFormCopyPasteControl2_ClipboardChanged()

Debug.Print "Changed 2 - " & VBFormCopyPasteControl2.IsFormDataInClipboard

End Sub
