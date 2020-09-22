VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAbout 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox RTB2 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5318
      _Version        =   393217
      BackColor       =   14737632
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmAbout.frx":0000
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long


Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
SetCurrentDirectory App.Path
RTB2.LoadFile ("About.txt")
End Sub
Private Sub Form_Resize()
RTB2.Height = Me.ScaleHeight
RTB2.Width = Me.ScaleWidth - 1800
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub
