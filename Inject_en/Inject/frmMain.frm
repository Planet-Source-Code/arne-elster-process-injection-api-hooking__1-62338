VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Process API Hook"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox lstMod 
      Height          =   3375
      Left            =   3300
      TabIndex        =   3
      Top             =   150
      Width           =   2790
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   315
      Left            =   150
      TabIndex        =   2
      Top             =   3600
      Width           =   1365
   End
   Begin VB.CommandButton cmdHook 
      Caption         =   "Inject"
      Height          =   315
      Left            =   1725
      TabIndex        =   1
      Top             =   3600
      Width           =   1365
   End
   Begin VB.ListBox lstProc 
      Height          =   3375
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Width           =   3090
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private procs() As PROCESSENTRY32

Private Sub ShowProcesses()
    Dim i   As Integer

    procs = GetProcesses()

    lstProc.Clear
    For i = LBound(procs) To UBound(procs) - 1
        lstProc.AddItem "  " & procs(i).szExeFile
    Next
End Sub

Private Sub cmdHook_Click()
    Dim hProcess    As Long

    hProcess = GetProcess(procs(lstProc.ListIndex).th32ProcessID)

    If Not InjectModule(hProcess, App.EXEName & ".exe") Then
        MsgBox "Could not inject module.", vbExclamation
        CloseProcess hProcess
        Exit Sub
    End If

    If Not RemoteHook(hProcess, "user32.dll", "MessageBoxA", AddressOf MessageBox) Then
        MsgBox "Couldn't hook MessageBoxA.", vbExclamation
        CloseProcess hProcess
        Exit Sub
    End If

    MsgBox "Successfully hooked.", vbInformation
    CloseProcess hProcess
End Sub

Private Sub cmdUpdate_Click()
    ShowProcesses
End Sub

Private Sub Form_Load()
    cmdUpdate_Click
End Sub

Private Sub lstProc_Click()
    Dim strModules()    As String
    Dim i               As Integer

    strModules = GetProcessModules(procs(lstProc.ListIndex).th32ProcessID)

    lstMod.Clear
    For i = 0 To UBound(strModules)
        If Not strModules(i) = "" Then
            lstMod.AddItem GetFile(strModules(i))
        End If
    Next
End Sub

Private Function GetFile(path As String) As String
    GetFile = Mid$(path, InStrRev(path, "\") + 1)
End Function
