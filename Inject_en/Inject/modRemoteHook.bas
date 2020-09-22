Attribute VB_Name = "modRemoteHook"
Option Explicit

Private Declare Function WriteProcessMemory Lib "kernel32" ( _
    ByVal ProcessHandle As Long, _
    lpBaseAddress As Any, _
    lpBuffer As Any, _
    ByVal nSize As Long, _
    lpNumberOfBytesWritten As Long _
) As Long

Private Declare Function ReadProcessMemory Lib "kernel32" ( _
    ByVal hProcess As Long, _
    lpBaseAddress As Any, _
    lpBuffer As Any, _
    ByVal nSize As Long, _
    lpNumberOfBytesWritten As Long _
) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
    ByVal lib As String _
) As String

Private Declare Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As Long, _
    ByVal lpProcName As String _
) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    source As Any, _
    ByVal Length As Long _
)

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
    ByVal lpModuleName As String _
) As Long

Private btOldAsm(4) As Byte

Public Function RemoteHook(ByVal hProcess As Long, _
                    ByVal module As String, _
                    ByVal fnc As String, _
                    ByVal NewAddr As Long) As Boolean
    
    Dim BytesWritten    As Long
    
    Dim hModule         As Long
    Dim hFnc            As Long

    Dim btNewAsm(4)     As Byte

    hModule = GetModuleHandle(module)
    If hModule = 0 Then
        hModule = LoadLibrary(module)
        If hModule = 0 Then Exit Function
    End If

    hFnc = GetProcAddress(hModule, fnc)
    If hFnc = 0 Then Exit Function

    ' save the first 4 bytes of the function to hook
    ReadProcessMemory hProcess, ByVal hFnc, btOldAsm(0), 5, BytesWritten
    If BytesWritten <> 5 Then Exit Function

    ' *** possible extension
    ' *** create a proxy function in the remote process
    ' *** to call the hooked function

    ' relative JMP address
    NewAddr = NewAddr - hFnc - 5

    btNewAsm(0) = &HE9                  ' JMP near
    CopyMemory btNewAsm(1), NewAddr, 4  ' rel. Addr

    ' overwrite function with JMP instruction
    WriteProcessMemory hProcess, ByVal hFnc, btNewAsm(0), 5, BytesWritten
    RemoteHook = BytesWritten = 5

End Function
