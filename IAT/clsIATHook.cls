VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIATHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***********************************
' Import Address Table Hook Class
'
' [rm_code]
'***********************************
' Thanks to:
' John Chamberlain's Compiler Controller
'***********************************

Private Const IMAGE_NUMBEROF_DIRECTIRY_ENRIES As Long = 16&
Private Const STANDARD_RIGHTS_REQUIRED        As Long = &HF0000
Private Const PAGE_EXECUTE_READWRITE          As Long = &H40&
Private Const IMAGE_NT_SIGNATURE              As Long = &H4550

Private Type IMAGE_FILE_HEADER
    Machine                 As Integer
    NumberOfSections        As Integer
    TimeDataStamp           As Long
    PointerToSymbolTable    As Long
    NumberOfSymbols         As Long
    SizeOfOptionalHeader    As Integer
    Characteristics         As Integer
End Type

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress          As Long
    Size                    As Long
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
    LoaderFlags             As Long
    NumberOfRvaAndSizes     As Long
    DataDirectory(IMAGE_NUMBEROF_DIRECTIRY_ENRIES - 1) As IMAGE_DATA_DIRECTORY
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

Private Declare Function GetProcAddress Lib "kernel32" ( _
            ByVal hModule As Long, _
            ByVal lpProcName As String) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
            ByVal lpModuleName As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            Destination As Any, _
            source As Any, _
            ByVal Length As Long)

Private Declare Function VirtualProtect Lib "kernel32" ( _
            lpAddress As Any, _
            ByVal dwSize As Long, _
            ByVal flNewProtect As Long, _
            lpflOldProtect As Long) As Long

Private lngOldAddr      As Long
Private lngIATPtr       As Long
Private blnHooked       As Boolean

Public Property Get OldAddress() As Long
    OldAddress = lngOldAddr
End Property

Public Property Get Hooked() As Boolean
    Hooked = blnHooked
End Property

Public Sub Unhook()
    If Not blnHooked Then Exit Sub

    ' restore old IAT entry
    PutMem lngIATPtr, lngOldAddr

    lngOldAddr = 0
    lngIATPtr = 0
    blnHooked = False
End Sub

Public Function Hook(ByVal module As String, _
                     ByVal lib As String, _
                     ByVal fnc As String, _
                     ByVal NewAddr As Long) As Long

    Dim hMod    As Long
    Dim hLib    As Long

    Dim DOSHdr  As IMAGE_DOS_HEADER
    Dim PEHdr   As IMAGE_OPTIONAL_HEADER32
    Dim ImgHdr  As IMAGE_DATA_DIRECTORY

    Dim lpIAT   As Long
    Dim IATLen  As Long
    Dim IATPos  As Long

    If blnHooked Then Exit Function

    hMod = GetModuleHandle(module)
    If hMod = 0 Then Exit Function

    hLib = GetModuleHandle(lib)
    If hLib = 0 Then Exit Function

    lngOldAddr = GetProcAddress(hLib, fnc)
    If lngOldAddr = 0 Then Exit Function

    ' read PE Header
    CopyMemory DOSHdr, ByVal hMod, LenB(DOSHdr)
    CopyMemory PEHdr, ByVal hMod + DOSHdr.e_lfanew, LenB(PEHdr)

    If PEHdr.Magic <> IMAGE_NT_SIGNATURE Then Exit Function

    ' IAT length and offset
    lpIAT = PEHdr.DataDirectory(15).VirtualAddress + hMod
    IATLen = PEHdr.DataDirectory(15).Size

    IATPos = lpIAT
    Do
        ' End of IAT?
        If IATPos > lpIAT + IATLen Then Exit Function

        ' found function?
        If DeRef(IATPos) = lngOldAddr Then
            ' overwrite pointer
            PutMem IATPos, NewAddr
            ' finished
            lngIATPtr = IATPos
            Hook = IATPos
            blnHooked = True
            Exit Function
        End If

        ' next function pointer
        IATPos = IATPos + 4
    Loop

End Function

Private Function DeRef(ByVal lpAddr As Long) As Long
    Dim lbuf As Long
    CopyMemory lbuf, ByVal lpAddr, 4
    DeRef = lbuf
End Function

Private Function PutMem(ByVal lpAddr As Long, ByVal lVal As Long) As Boolean
    Dim lngOldProtect   As Long

    If 0 = VirtualProtect(ByVal lpAddr, 4, PAGE_EXECUTE_READWRITE, lngOldProtect) Then
        Exit Function
    End If

    CopyMemory ByVal lpAddr, lVal, 4
    VirtualProtect ByVal lpAddr, 4, lngOldProtect, lngOldProtect

    PutMem = True
End Function
