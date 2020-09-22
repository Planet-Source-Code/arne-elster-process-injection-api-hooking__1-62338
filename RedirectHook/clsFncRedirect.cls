VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFncRedirect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***********************************
' function redirection class
'
' [rm_code]
'***********************************
' Thanks to:
' EBArtSoft's API HOOK Demo II
'***********************************

Private Declare Function VirtualProtect Lib "kernel32" ( _
            lpAddress As Any, _
            ByVal dwSize As Long, _
            ByVal flNewProtect As Long, _
            lpflOldProtect As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            pDest As Any, _
            pSource As Any, _
            ByVal dwLength As Long)

Private Declare Function GetProcAddress Lib "kernel32" ( _
            ByVal hModule As Long, _
            ByVal lpProcName As String) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
            ByVal lpModuleName As String) As Long

Private Const PAGE_EXECUTE_READWRITE          As Long = &H40&

Private blnHooked       As Boolean      ' function hooked?
Private lpOldAddr       As Long         ' address of hooked function

Private btOldASM(4)     As Byte         ' old 5 bytes of hooked function
Private btReal(31)      As Byte         ' hooked function

' restore old hooked function
Public Function Unhook() As Boolean
    If Not blnHooked Then Exit Function

    ' overwrite new with the old instruction
    blnHooked = PutMem(lpOldAddr, VarPtr(btOldASM(0)), UBound(btOldASM) + 1)
    Unhook = blnHooked
End Function

Public Property Get OldAddress() As Long
    OldAddress = lpOldAddr
End Property

Public Property Get Hooked() As Boolean
    Hooked = blnHooked
End Property

' redirect a exported function of a module to an other one
'
'   Param1: exporting module (eg "kernel32")
'   Param2: target function (eg "Sleep")
'   Param3: address of new function
' [Param4]: address of a function which
'           will point to new old one
'
Public Function Hook(ByVal module As String, _
        ByVal fnc As String, _
        ByVal NewAddr As Long, _
        Optional ProxyAddr As Long) As Boolean

    Dim hModule     As Long
    Dim hFnc        As Long

    If blnHooked Then Exit Function

    hModule = GetModuleHandle(module)
    If hModule = 0 Then Exit Function

    hFnc = GetProcAddress(hModule, fnc)
    If hFnc = 0 Then Exit Function
    lpOldAddr = hFnc

    ' save old instructions
    If Not GetMem(hFnc, VarPtr(btOldASM(0)), UBound(btOldASM) + 1) Then
        Exit Function
    End If

    ' redirect ProxyAddr to target function
    If ProxyAddr <> 0 Then
        CopyMemory btReal(0), btOldASM(0), UBound(btOldASM) + 1
        Redirect VarPtr(btReal(UBound(btOldASM) + 1)), lpOldAddr + UBound(btOldASM) + 1
        Redirect ProxyAddr, VarPtr(btReal(0))
    End If

    ' redirect the target function to the replacement function
    blnHooked = Redirect(hFnc, NewAddr)
    Hook = blnHooked

End Function

' write a JMP near instruction to an address
Public Function Redirect(ByVal OldAddr As Long, ByVal NewAddr As Long) As Boolean
    Dim btAsm(4)    As Byte
    Dim lngNewAddr  As Long

    ' relative jump address
    lngNewAddr = NewAddr - OldAddr - (UBound(btAsm) + 1)

    btAsm(0) = &HE9                     ' JMP near
    CopyMemory btAsm(1), lngNewAddr, 4  ' rel. addr

    Redirect = PutMem(OldAddr, VarPtr(btAsm(0)), UBound(btAsm) + 1)

End Function

Private Function GetMem(ByVal lpAddr As Long, ByVal pData As Long, ByVal dlen As Long) As Boolean
    Dim lngOldProtect   As Long

    If 0 = VirtualProtect(ByVal lpAddr, dlen, PAGE_EXECUTE_READWRITE, lngOldProtect) Then
        Exit Function
    End If

    CopyMemory ByVal pData, ByVal lpAddr, dlen
    VirtualProtect ByVal lpAddr, dlen, lngOldProtect, lngOldProtect

    GetMem = True
End Function

Private Function PutMem(ByVal lpAddr As Long, ByVal pData As Long, ByVal dlen As Long) As Boolean
    Dim lngOldProtect   As Long

    If 0 = VirtualProtect(ByVal lpAddr, dlen, PAGE_EXECUTE_READWRITE, lngOldProtect) Then
        Exit Function
    End If

    CopyMemory ByVal lpAddr, ByVal pData, dlen
    VirtualProtect ByVal lpAddr, dlen, lngOldProtect, lngOldProtect

    PutMem = True
End Function
