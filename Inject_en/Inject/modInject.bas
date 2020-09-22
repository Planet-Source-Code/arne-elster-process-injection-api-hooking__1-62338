Attribute VB_Name = "modInject"
Option Explicit

Public Const IMAGE_NT_SIGNATURE               As Long = &H4550

Private Const PAGE_EXECUTE_READWRITE          As Long = &H40&
Private Const PROCESS_ALL_ACCESS              As Long = &H1F0FFF
Private Const MEM_COMMIT                      As Long = &H1000&
Private Const MEM_RESERVE                     As Long = &H2000&

Private Const TH32CS_SNAPPROCESS              As Long = 2&
Private Const MAX_PATH                        As Long = 260&

Public Type PROCESSENTRY32
  dwSize                    As Long
  cntUsage                  As Long
  th32ProcessID             As Long
  th32DefaultHeapID         As Long
  th32ModuleID              As Long
  cntThreads                As Long
  th32ParentProcessID       As Long
  pcPriClassBase            As Long
  dwFlags                   As Long
  szExeFile                 As String * MAX_PATH
End Type

Private Type IMAGE_FILE_HEADER
    Machine                 As Integer
    NumberOfSections        As Integer
    TimeDataStamp           As Long
    PointerToSymbolTable    As Long
    NumberOfSymbols         As Long
    SizeOfOptionalHeader    As Integer
    Characteristics         As Integer
End Type

Private Type IMAGE_OPTIONAL_HEADER32
    Magic                   As Integer
    MajorLinkerVersion      As Byte
    MinorLinkerVersion      As Byte
    SizeOfCode              As Long
    SizeOfInitalizedData    As Long
    SizeOfUninitalizedData  As Long
    AddressOfEntryPoint     As Long
    BaseOfCode              As Long
    BaseOfData              As Long
    ImageBase               As Long
    SectionAlignment        As Long
    FileAlignment           As Long
    MajorOperatingSystemVer As Integer
    MinorOperatingSystemVer As Integer
    MajorImageVersion       As Integer
    MinorImageVersion       As Integer
    MajorSubsystemVersion   As Integer
    MinorSubsystemVersion   As Integer
    Reserved1               As Long
    SizeOfImage             As Long
    SizeOfHeaders           As Long
    CheckSum                As Long
    Subsystem               As Integer
    DllCharacteristics      As Integer
    SizeOfStackReserve      As Long
    SizeOfStackCommit       As Long
    SizeOfHeapReserve       As Long
    SizeOfHeapCommit        As Long
    LoaerFlags              As Long
    NumberOfRvaAndSizes     As Long
End Type

Private Type IMAGE_DOS_HEADER
    e_magic                 As Integer
    e_cblp                  As Integer
    e_cp                    As Integer
    e_crlc                  As Integer
    e_cparhdr               As Integer
    e_minalloc              As Integer
    e_maxalloc              As Integer
    e_ss                    As Integer
    e_sp                    As Integer
    e_csum                  As Integer
    e_ip                    As Integer
    e_cs                    As Integer
    e_lfarlc                As Integer
    e_onvo                  As Integer
    e_res(3)                As Integer
    e_oemid                 As Integer
    e_oeminfo               As Integer
    e_res2(9)               As Integer
    e_lfanew                As Long
End Type

Private Declare Function EnumProcessModules Lib "psapi.dll" ( _
    ByVal hProcess As Long, ByRef lphModule As Long, _
    ByVal cb As Long, ByRef lpcbNeeded As Long _
) As Long

Private Declare Function GetModuleFileName Lib "kernel32.dll" _
Alias "GetModuleFileNameA" ( _
    ByVal hModule As Long, ByVal lpFileName As String, _
    ByVal nSize As Long _
) As Long

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
    ByVal lFlags As Long, ByVal lProcessID As Long _
) As Long

Private Declare Function ProcessFirst Lib "kernel32" _
Alias "Process32First" ( _
    ByVal hSnapShot As Long, uProcess As PROCESSENTRY32 _
) As Long

Private Declare Function ProcessNext Lib "kernel32" _
Alias "Process32Next" ( _
    ByVal hSnapShot As Long, uProcess As PROCESSENTRY32 _
) As Long

Private Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long _
) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long _
) As Long

Private Declare Function VirtualAllocEx Lib "kernel32" ( _
    ByVal ProcessHandle As Long, _
    ByVal lpAddress As Long, _
    ByVal dwSize As Long, _
    ByVal flAllocationType As Long, _
    ByVal flProtect As Long _
) As Long

Private Declare Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As Long, _
    ByVal lpProcName As String _
) As Long

Private Declare Function WriteProcessMemory Lib "kernel32" ( _
    ByVal ProcessHandle As Long, _
    lpBaseAddress As Any, _
    lpBuffer As Any, _
    ByVal nSize As Long, _
    lpNumberOfBytesWritten As Long _
) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    source As Any, _
    ByVal Length As Long _
)

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
    ByVal lpModuleName As String _
) As Long

Public Function GetProcess(ByVal pid As Long) As Long
    GetProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, pid)
End Function

Public Sub CloseProcess(hProcess As Long)
    CloseHandle hProcess: hProcess = 0
End Sub

Public Function InjectModule(ByVal hProcess As Long, ByVal module As String) As Boolean
    Dim DOSHdr  As IMAGE_DOS_HEADER
    Dim PEHdr   As IMAGE_OPTIONAL_HEADER32
    Dim ImgHdr  As IMAGE_FILE_HEADER

    Dim hModule As Long, hNewModule     As Long
    Dim ModSize As Long, BytesWritten   As Long
    Dim TID             As Long

    hModule = GetModuleHandle(module)
    If hModule = 0 Then Exit Function

    ' read PE Header
    CopyMemory DOSHdr, ByVal hModule, LenB(DOSHdr)
    CopyMemory PEHdr, ByVal hModule + DOSHdr.e_lfanew + 4 + Len(ImgHdr), LenB(PEHdr)

    ' size of module in memory
    ModSize = PEHdr.SizeOfImage

    ' alloc some space in the target process
    hNewModule = VirtualAllocEx(hProcess, hModule, ModSize, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If hNewModule = 0 Then Exit Function

    ' copy module to target process
    WriteProcessMemory hProcess, ByVal hNewModule, ByVal hModule, ModSize, BytesWritten
    InjectModule = CBool(BytesWritten = ModSize)
End Function

Public Function GetProcesses() As PROCESSENTRY32()
    Dim hSnapShot As Long, Result As Long
    Dim aa      As String, bb   As String

    Dim Process         As PROCESSENTRY32
    Dim processes()     As PROCESSENTRY32
    ReDim processes(0) As PROCESSENTRY32

    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapShot = 0 Then Exit Function

    Process.dwSize = Len(Process)
    Result = ProcessFirst(hSnapShot, Process)

    Do While Result <> 0

        With Process
            aa = Left$(.szExeFile, InStr(.szExeFile, Chr$(0)) - 1)
        End With

        If Right$(LCase(aa), 3) = "exe" Then
            processes(UBound(processes)) = Process
            ReDim Preserve processes(UBound(processes) + 1)
        End If

        Result = ProcessNext(hSnapShot, Process)

    Loop

    CloseHandle hSnapShot
    GetProcesses = processes
End Function

Public Function GetProcessModules(pid As Long) As String()

    Dim nProcID             As Long
    Dim i                   As Long

    Dim nResult             As Long
    Dim nTemp               As Long
    Dim lModules(1 To 200)  As Long

    Dim strProcess          As String

    Dim hProcess            As Long
    Dim processes()         As PROCESSENTRY32

    Dim strModFile          As String * 255
    Dim strModFiles()       As String
    ReDim strModFiles(0) As String

    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, pid)
    nResult = EnumProcessModules(hProcess, lModules(1), 200, nTemp)

    For i = 1 To UBound(lModules)
        If lModules(i) <> 0 Then
            strModFiles(UBound(strModFiles)) = Left$(strModFile, GetModuleFileName(lModules(i), strModFile, Len(strModFile)))
            ReDim Preserve strModFiles(UBound(strModFiles) + 1)
        End If
    Next

    CloseHandle hProcess

    GetProcessModules = strModFiles

End Function
