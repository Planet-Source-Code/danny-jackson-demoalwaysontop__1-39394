VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Always On Top"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   6530
      TabIndex        =   6
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdForm4 
      Caption         =   "Form 4"
      Height          =   375
      Left            =   5260
      TabIndex        =   5
      ToolTipText     =   "Run Form 4"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdForm3 
      Caption         =   "Form 3"
      Height          =   375
      Left            =   4000
      TabIndex        =   4
      ToolTipText     =   "Run Form 3"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdForm2 
      Caption         =   "Form 2"
      Height          =   375
      Left            =   2740
      TabIndex        =   3
      ToolTipText     =   "Run Form 2"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdForm1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Form 1"
      Height          =   375
      Left            =   1450
      TabIndex        =   2
      ToolTipText     =   "Run Form1"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   195
      TabIndex        =   1
      ToolTipText     =   "Exit the program"
      Top             =   5160
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   4455
      Left            =   195
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7858
      _Version        =   393217
      BackColor       =   14737632
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

Private Sub cmdAbout_Click()
frmAbout.Show
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdForm1_Click()
frmForm1.Show
End Sub

Private Sub cmdForm2_Click()
frmForm2.Show
End Sub

Private Sub cmdForm3_Click()
frmForm3.Show
End Sub

Private Sub cmdForm4_Click()
frmForm4.Show
End Sub

Private Sub Form_Resize()
RTB1.Height = Me.ScaleHeight - 800
RTB1.Width = Me.ScaleWidth - 300
cmdExit.Top = RTB1.Height + 300
cmdForm1.Top = cmdExit.Top
cmdForm2.Top = cmdForm1.Top
cmdForm3.Top = cmdForm2.Top
cmdForm4.Top = cmdForm3.Top
cmdAbout.Top = cmdForm4.Top
End Sub

Private Sub Form_Load()
SetCurrentDirectory App.Path
RTB1.LoadFile ("AlwaysOnTop.txt")
End Sub
