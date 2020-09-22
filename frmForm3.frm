VERSION 5.00
Begin VB.Form frmForm3 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Text Form 3"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton cmdClose3 
      Caption         =   "Close"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Close Form 3"
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "frmForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const SWP_SHOWWINDOW = &H40


Private Sub cmdClose3_Click()
Unload Me
End Sub

Private Sub Form_Load()

  Dim TempValue As Long
  Dim MyWidth As Long, MyHeight As Long
  Dim MyTop As Long, MyLeft As Long
  
  
  MyWidth = (Screen.Width / 4)
  MyWidth = MyWidth / Screen.TwipsPerPixelX
  MyHeight = Screen.Height / 4
  MyHeight = MyHeight / Screen.TwipsPerPixelY
  
  MyLeft = Screen.Width / (2 * Screen.TwipsPerPixelX)
  MyTop = Screen.Height / (2 * Screen.TwipsPerPixelY)
'Call SetWindowPos API function
  TempValue = SetWindowPos(Me.hwnd, HWND_TOPMOST, MyLeft, MyTop, MyWidth, MyHeight, SWP_SHOWWINDOW)
  Text3.Text = "My Settings are:" & vbCrLf & "MyWidth = (Screen.Width / 4)" & vbCrLf & "MyHeight = Screen.Height / 4" & vbCrLf & "MyLeft = Screen.Width / (2 * Screen.TwipsPerPixelX)" & vbCrLf & "MyTop = Screen.Height / (2 * Screen.TwipsPerPixelY)"
End Sub

