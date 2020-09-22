Attribute VB_Name = "modCallFnc"
Option Explicit

' modified version of:
' http://nienie.com/~masapico/doc_FuncPtr.html
'
' call function pointers in new threads

Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal Flags As Long, ByVal Size As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal Mem As Long) As Long
Private Declare Function MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Src As Any, ByVal Size As Long) As Long

Private Const GMEM_FIXED As Long = 0&

Public Function CallFuncPtr(FuncPtr As Long, ParamArray Params() As Variant) As Long
    Const MAX_CODESIZE  As Long = 65536

    Dim I               As Long, pCodeData      As Long
    Dim pParamData()    As Long, PC             As Long
    Dim Operand         As Long, RetValue       As Long
    Dim LongValue       As Long, dwThreadID     As Long
    Dim hThread         As Long, dwExit         As Long
    Dim StrValue        As String

    ReDim pParamData(UBound(Params)) As Long
    pCodeData = GlobalAlloc(GMEM_FIXED, MAX_CODESIZE)
    PC = pCodeData

    AddByte PC, &H55

    For I = UBound(Params) To 0 Step -1
        If VarType(Params(I)) = vbString Then
            pParamData(I) = GlobalAlloc(GMEM_FIXED, _
                                LenB(Params(I)))
            StrValue = Params(I)
            MoveMemory ByVal pParamData(I), _
                       ByVal StrValue, LenB(StrValue)
            Operand = pParamData(I)
        Else
            Operand = Params(I)
        End If
        AddByte PC, &H68
        AddLong PC, Operand
    Next

    AddByte PC, &HB8
    AddLong PC, FuncPtr
    AddInt PC, &HD0FF
    AddByte PC, &HBA
    AddLong PC, VarPtr(RetValue)
    AddInt PC, &H289
    AddByte PC, &H5D
    AddInt PC, &HC033
    AddByte PC, &HC2
    AddInt PC, &H8

    hThread = CreateThread(0, 0, pCodeData, _
                           0, 0, dwThreadID)

    Do
        GetExitCodeThread hThread, dwExit
        If dwExit <> 259 Then Exit Do
        DoEvents
    Loop

    GlobalFree pCodeData
    For I = 0 To UBound(Params)
        If pParamData(I) <> 0 Then
            GlobalFree pParamData(I)
        End If
    Next

    CallFuncPtr = RetValue
End Function

Private Sub AddByte(ByRef PC As Long, ByVal ByteValue As Byte)
    MoveMemory ByVal PC, ByteValue, 1
    PC = PC + 1
End Sub

Private Sub AddInt(ByRef PC As Long, ByVal IntValue As Integer)
    MoveMemory ByVal PC, IntValue, 2
    PC = PC + 2
End Sub

Private Sub AddLong(ByRef PC As Long, ByVal LongValue As Long)
    MoveMemory ByVal PC, LongValue, 4
    PC = PC + 4
End Sub
