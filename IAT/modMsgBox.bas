Attribute VB_Name = "modMsgBox"
Option Explicit

Public clsGPAHook   As clsIATHook

Private lpGetExitCodeThread As Long
Private lpRtlMoveMemory     As Long
Private lpCreateThread      As Long
Private lpGlobalAlloc       As Long
Private lpGlobalFree        As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

' avoid deadlock load function pointers needed
' by modCallFnc
Public Sub PrepareGPAHook()
    Dim hKernel As Long
    hKernel = GetModuleHandle("kernel32")
    lpGetExitCodeThread = GetProcAddress(hKernel, "GetExitCodeThread")
    lpRtlMoveMemory = GetProcAddress(hKernel, "RtlMoveMemory")
    lpCreateThread = GetProcAddress(hKernel, "CreateThread")
    lpGlobalAlloc = GetProcAddress(hKernel, "GlobalAlloc")
    lpGlobalFree = GetProcAddress(hKernel, "GlobalFree")
End Sub

Public Function GetProcAddressEx(ByVal hModule As Long, _
                                 ByVal szFnc As String) As Long

    Dim btFnc() As Byte

    szFnc = StrConv(szFnc, vbUnicode)
    btFnc = StrConv(szFnc, vbFromUnicode)

    ' remove null char
    szFnc = Left$(szFnc, Len(szFnc) - 1)

    Select Case szFnc
        Case "MessageBoxA"
            ' redirect to new MessageBoxA
            GetProcAddressEx = GetAdr(AddressOf MessageBoxAEx)
        Case "GetExitCodeThread"
            GetProcAddressEx = lpGetExitCodeThread
        Case "RtlMoveMemory"
            GetProcAddressEx = lpRtlMoveMemory
        Case "CreateThread"
            GetProcAddressEx = lpCreateThread
        Case "GlobalAlloc"
            GetProcAddressEx = lpGlobalAlloc
        Case "GlobalFree"
            GetProcAddressEx = lpGlobalFree
        Case Else
            ' unknown function, call old GPA
            GetProcAddressEx = CallFuncPtr(clsGPAHook.OldAddress, hModule, btFnc(0))
    End Select

End Function

Public Function MessageBoxAEx(ByVal hWnd As Long, _
                              ByVal szText As String, _
                              ByVal szTitle As String, _
                              ByVal dwStyle As Long) As Long

    szText = StrConv(szText, vbUnicode)
    szTitle = "Powered by IAT Hook"
    MessageBoxAEx = MsgBox(szText, dwStyle, szTitle)

End Function

Private Function GetAdr(lng As Long) As Long
    GetAdr = lng
End Function
