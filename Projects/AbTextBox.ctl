VERSION 5.00
Begin VB.UserControl AbTextBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1740
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   1740
End
Attribute VB_Name = "AbTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=========================================================================
'
'   AbTextBox Project
'   Copyright (c) 2006 PeterJack
'
'  Textbox with Arabic support and RTL( right to left ).
'  other source code:
'  http://www.AmazeBrowser.com/sourcecode.htm
'  Require win2k,winxp
'=========================================================================

Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long) As Long
Private Declare Function GetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLengthW Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private hWndEdit As Long
Private m_sText As String

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If hWndEdit = 0 Then
        hWndEdit = CreateWindowExW(768, ByVal StrPtr("EDIT"), ByVal StrPtr(""), 1342177664, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hwnd, 0, App.hInstance, ByVal 0&)
        Dim fnt As IFont

        Set fnt = UserControl.Font
        SendMessageW hWndEdit, 48, fnt.hFont, ByVal 1&
        Set fnt = Nothing

    End If
End Sub

Public Property Get Text() As String
Dim nLen As Long

    nLen = GetWindowTextLengthW(hWndEdit)
    
    If (nLen > 0) Then
        m_sText = String$(nLen + 1, 0)
        GetWindowTextW hWndEdit, StrPtr(m_sText), nLen
    
        nLen = InStr(m_sText, vbNullChar)
        If nLen > 1 Then
            Text = Left$(m_sText, nLen - 1)
        Else
            Text = m_sText
        End If
    End If
End Property

Public Property Let Text(ByVal sText As String)

    m_sText = sText
    If hWndEdit <> 0 Then
        SetWindowTextW hWndEdit, ByVal StrPtr(sText)
    End If
    Call UserControl.PropertyChanged("Text")

End Property
