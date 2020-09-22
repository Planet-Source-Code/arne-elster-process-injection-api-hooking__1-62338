VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Funktionsumleitung"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   113
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdRemHook 
      Caption         =   "Remove hook"
      Height          =   465
      Left            =   825
      TabIndex        =   1
      Top             =   900
      Width           =   2865
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Exit with End"
      Height          =   465
      Left            =   825
      TabIndex        =   0
      Top             =   300
      Width           =   2865
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetProcAddress Lib "kernel32" ( _
            ByVal hModule As Long, _
            ByVal lpProcName As String) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
            ByVal lpModuleName As String) As Long

Private clsRedirect As clsFncRedirect

Private Sub cmdEnd_Click()
    ' call EndEx
    End
End Sub

Private Sub cmdRemHook_Click()
    If clsRedirect.Unhook Then
        MsgBox "Function restored.", vbInformation
    Else
        MsgBox "failed.", vbExclamation
    End If
End Sub

Private Sub Form_Load()
    Set clsRedirect = New clsFncRedirect

    If GetModuleHandle("vba6") Then
        MsgBox "Run only compiled!", vbExclamation
        End
    End If

    ' redirect __vbaEnd to EndEx,
    ' make RealEnd proxy to __vbaEnd
    If clsRedirect.Hook("msvbvm60", "__vbaEnd", AddressOf EndEx, AddressOf RealEnd) Then
        MsgBox "End got hooked!", vbInformation
    Else
        MsgBox "Hook failed", vbExclamation
    End If

End Sub
