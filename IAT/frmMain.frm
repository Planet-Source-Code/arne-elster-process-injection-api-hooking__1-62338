VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "VB IAT Hook"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   56
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdShow 
      Caption         =   "MessageBoxA"
      Height          =   465
      Left            =   975
      TabIndex        =   0
      Top             =   225
      Width           =   2715
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MessageBoxA Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal strText As String, _
    ByVal strTitle As String, _
    ByVal dwStyle As Long) As Long

Private Sub cmdShow_Click()
    ' VB will first call GetProcAddress,
    ' and then MessageBoxA
    MessageBoxA 0, "Text", "Titel", vbInformation
End Sub

Private Sub Form_Load()
    Set clsGPAHook = New clsIATHook

    If GetModuleHandle("vba6") Then
        MsgBox "Run only compiled!", vbExclamation
        End
    End If

    ' prepare function pointers used by CallFuncPtr
    PrepareGPAHook

    ' hook GetProcAddress imported by the VB runtime
    If clsGPAHook.Hook("msvbvm60", "kernel32", "GetProcAddress", AddressOf GetProcAddressEx) = 0 Then
        MsgBox "could not hook", vbExclamation
    Else
        MsgBox "GetProcAddress successfully hooked.", vbInformation
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    clsGPAHook.Unhook
End Sub
