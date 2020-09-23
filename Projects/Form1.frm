VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "AbTextBox - Arabic font"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   11.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "VB Msgbox"
      Height          =   330
      Left            =   3825
      TabIndex        =   7
      Top             =   2565
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   90
      TabIndex        =   5
      Top             =   1260
      Width           =   5100
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Our Msgbox"
      Height          =   330
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1590
   End
   Begin Project1.AbTextBox AbTextBox1 
      Height          =   420
      Left            =   90
      TabIndex        =   0
      Top             =   405
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   741
   End
   Begin VB.Label Label4 
      Caption         =   "Arabic char in VB Textbox:"
      Height          =   330
      Left            =   90
      TabIndex        =   6
      Top             =   900
      Width           =   5100
   End
   Begin VB.Label Label3 
      Caption         =   "Email: peterjack2006@gmail.com"
      ForeColor       =   &H00C0C0C0&
      Height          =   330
      Left            =   360
      TabIndex        =   4
      Top             =   2610
      Width           =   4830
   End
   Begin VB.Label Label2 
      Caption         =   "Textbox with Arabic support and RTL( right to left )"
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   5100
   End
   Begin VB.Label Label1 
      Caption         =   "http://www.AmazeBrowser.com/"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   2295
      Width           =   3435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long

Private Sub Command1_Click()
    Dim sT As String
    sT = AbTextBox1.Text
    
    MessageBoxW Me.hwnd, StrPtr(sT), StrPtr(""), 0&

End Sub

Private Sub Command2_Click()
    Dim sT As String
    sT = AbTextBox1.Text
    
    MsgBox sT
    
End Sub

Private Sub Form_Load()
    Dim vS As Variant
    vS = Array(32, 1740, 1593, 1606, 1740, 32, 1576, 1575, 32, 1601, 1585, 1587, 1578, 1575, 1583, 1606)
    Dim sT As String
    Dim i As Long
    
    For i = LBound(vS) To UBound(vS)
        sT = sT & ChrW(vS(i))
    Next i
    
    'our AbTextBox
    AbTextBox1.Text = sT
    'vb TextBox
    Text1.Text = sT
End Sub

Private Sub Label1_Click()
    MsgBox "if you want to see a demo, just go to my homepage, and try AmazeCopy" & vbCrLf & "http://www.AmazeBrowser.com", vbInformation
End Sub
