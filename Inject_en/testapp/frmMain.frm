VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "API Hook Target"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "MsgBox"
      Height          =   465
      Left            =   750
      TabIndex        =   1
      Top             =   1050
      Width           =   3390
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MessageBoxA"
      Height          =   465
      Left            =   750
      TabIndex        =   0
      Top             =   525
      Width           =   3390
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MessageBoxA Lib "user32" ( _
    ByVal hwnd As Long, ByVal Msg As String, ByVal title As String, ByVal style As Long _
) As Long

Private Sub Command1_Click()
    If MessageBoxA(0, "prompt", "title", vbQuestion Or vbYesNo) = vbYes Then
        MsgBox "Yes"
    Else
        MsgBox "No"
    End If
End Sub

Private Sub Command2_Click()
    If MsgBox("prompt", vbQuestion Or vbYesNo, "title") = vbYes Then
        MsgBox "Yes"
    Else
        MsgBox "No"
    End If
End Sub
