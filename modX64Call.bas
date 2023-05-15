Attribute VB_Name = "modX64Call"

' //
' // modX64Call.bas
' // Module for calling functions in long-mode (x64)
' // by The trick 2018 - 2023
' // v.1.0.1
' //

Option Explicit

Private Const ProcessBasicInformation         As Long = 0
Private Const MEM_RESERVE                     As Long = &H2000&
Private Const MEM_COMMIT                      As Long = &H1000&
Private Const MEM_RELEASE                     As Long = &H8000&
Private Const PAGE_READWRITE                  As Long = 4&
Private Const PAGE_READONLY                   As Long = 2&
Private Const PAGE_EXECUTE_READ               As Long = &H20&
Private Const FADF_AUTO                       As Long = 1
Private Const PAGE_EXECUTE_READWRITE          As Long = &H40&
Private Const PROCESS_VM_READ                 As Long = &H10
Private Const LOAD_LIBRARY_AS_IMAGE_RESOURCE  As Long = &H20
Private Const OPEN_EXISTING                   As Long = 3
Private Const GENERIC_READ                    As Long = &H80000000
Private Const GENERIC_EXECUTE                 As Long = &H20000000
Private Const SEC_IMAGE                       As Long = &H1000000
Private Const FILE_ATTRIBUTE_NORMAL           As Long = &H80
Private Const MAX_PATH                        As Long = 260
Private Const INVALID_HANDLE_VALUE            As Long = -1
Private Const FILE_MAP_READ                   As Long = 4
Private Const FILE_MAP_EXECUTE                As Long = &H20
Private Const PROCESSOR_ARCHITECTURE_AMD64    As Long = 9

Private Type UNICODE_STRING64
    Length                          As Integer
    MaxLength                       As Integer
    lPad                            As Long
    lpBuffer                        As Currency
End Type

Private Type ANSI_STRING64
    Length                          As Integer
    MaxLength                       As Integer
    lPad                            As Long
    lpBuffer                        As Currency
End Type

Private Type PROCESS_BASIC_INFORMATION64
    ExitStatus                      As Long
    Reserved0                       As Long
    PebBaseAddress                  As Currency
    AffinityMask                    As Currency
    BasePriority                    As Long
    Reserved1                       As Long
    uUniqueProcessId                As Currency
    uInheritedFromUniqueProcessId   As Currency
End Type

Private Type IMAGE_FILE_HEADER
    Machine                         As Integer
    NumberOfSections                As Integer
    TimeDateStamp                   As Long
    PointerToSymbolTable            As Long
    NumberOfSymbols                 As Long
    SizeOfOptionalHeader            As Integer
    Characteristics                 As Integer
End Type

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress                  As Long
    Size                            As Long
End Type

Private Type IMAGE_EXPORT_DIRECTORY
    Characteristics                 As Long
    TimeDateStamp                   As Long
    MajorVersion                    As Integer
    MinorVersion                    As Integer
    pName                           As Long
    Base                            As Long
    NumberOfFunctions               As Long
    NumberOfNames                   As Long
    AddressOfFunctions              As Long
    AddressOfNames                  As Long
    AddressOfNameOrdinals           As Long
End Type

Private Type IMAGE_SECTION_HEADER
    SectionName(1)                  As Long
    VirtualSize                     As Long
    VirtualAddress                  As Long
    SizeOfRawData                   As Long
    PointerToRawData                As Long
    PointerToRelocations            As Long
    PointerToLinenumbers            As Long
    NumberOfRelocations             As Integer
    NumberOfLinenumbers             As Integer
    Characteristics                 As Long
End Type

Private Type SAFEARRAYBOUND
    cElements                       As Long
    lLbound                         As Long
End Type

Private Type SAFEARRAY1D
    cDims                           As Integer
    fFeatures                       As Integer
    cbElements                      As Long
    cLocks                          As Long
    pvData                          As Long
    Bounds                          As SAFEARRAYBOUND
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize             As Long
    dwMajorVersion                  As Long
    dwMinorVersion                  As Long
    dwBuildNumber                   As Long
    dwPlatformId                    As Long
    szCSDVersion                    As String * 128
End Type

Private Type VS_FIXEDFILEINFO
    dwSignature                     As Long
    dwStrucVersion                  As Long
    dwFileVersionMS                 As Long
    dwFileVersionLS                 As Long
    dwProductVersionMS              As Long
    dwProductVersionLS              As Long
    dwFileFlagsMask                 As Long
    dwFileFlags                     As Long
    dwFileOS                        As Long
    dwFileType                      As Long
    dwFileSubtype                   As Long
    dwFileDateMS                    As Long
    dwFileDateLS                    As Long
End Type

Private Type SYSTEM_INFO
    wProcessorArchitecture          As Integer
    wReserved                       As Integer
    dwPageSize                      As Long
    lpMinimumApplicationAddress     As Long
    lpMaximumApplicationAddress     As Long
    dwActiveProcessorMask           As Long
    dwNumberOrfProcessors           As Long
    dwProcessorType                 As Long
    dwAllocationGranularity         As Long
    dwReserved                      As Long
End Type

Private Declare Sub GetNativeSystemInfo Lib "kernel32" ( _
                    ByRef lpSystemInfo As SYSTEM_INFO)
Private Declare Function OpenProcess Lib "kernel32" ( _
                         ByVal dwDesiredAccess As Long, _
                         ByVal bInheritHandle As Long, _
                         ByVal dwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long
Private Declare Function NtWow64QueryInformationProcess64 Lib "ntdll" ( _
                         ByVal hProcess As Long, _
                         ByVal ProcessInformationClass As Long, _
                         ByRef pProcessInformation As Any, _
                         ByVal uProcessInformationLength As Long, _
                         ByRef puReturnLength As Long) As Long
Private Declare Function NtWow64ReadVirtualMemory64 Lib "ntdll" ( _
                         ByVal hProcess As Long, _
                         ByVal p64Address As Currency, _
                         ByRef Buffer As Any, _
                         ByVal l64BufferLen As Currency, _
                         ByRef pl64ReturnLength As Currency) As Long
Private Declare Function GetMem8 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef Dst As Any) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef Dst As Any) As Long
Private Declare Function PutMem4 Lib "msvbvm60" ( _
                         ByRef pDst As Any, _
                         ByVal lVal As Long) As Long
Private Declare Function GetMem2 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef Dst As Any) As Long
Private Declare Function GetMem1 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef Dst As Any) As Long
Private Declare Function VirtualAlloc Lib "kernel32" ( _
                         ByVal lpAddress As Long, _
                         ByVal dwSize As Long, _
                         ByVal flAllocationType As Long, _
                         ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" ( _
                         ByVal lpAddress As Long, _
                         ByVal dwSize As Long, _
                         ByVal dwFreeType As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32.dll" ( _
                         ByRef pvInstance As Any, _
                         ByVal oVft As Long, _
                         ByVal cc As Long, _
                         ByVal vtReturn As VbVarType, _
                         ByVal cActuals As Long, _
                         ByRef prgvt As Any, _
                         ByRef prgpvarg As Any, _
                         ByRef pvargResult As Variant) As Long
Private Declare Function lstrcmp Lib "kernel32" _
                         Alias "lstrcmpA" ( _
                         ByRef lpString1 As Any, _
                         ByRef lpString2 As Any) As Long
Private Declare Function lstrcmpi Lib "kernel32" _
                         Alias "lstrcmpiA" ( _
                         ByRef lpString1 As Any, _
                         ByRef lpString2 As Any) As Long
Private Declare Function ArrPtr Lib "msvbvm60" _
                         Alias "VarPtr" ( _
                         ByRef psa() As Any) As Long
Private Declare Function Wow64DisableWow64FsRedirection Lib "kernel32" ( _
                         ByRef lvalue As Long) As Long
Private Declare Function Wow64RevertWow64FsRedirection Lib "kernel32" ( _
                         ByVal lvalue As Long) As Long
Private Declare Function CreateFile Lib "kernel32" _
                         Alias "CreateFileW" ( _
                         ByVal lpFileName As Long, _
                         ByVal dwDesiredAccess As Long, _
                         ByVal dwShareMode As Long, _
                         ByRef lpSecurityAttributes As Any, _
                         ByVal dwCreationDisposition As Long, _
                         ByVal dwFlagsAndAttributes As Long, _
                         ByVal hTemplateFile As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" _
                         Alias "GetSystemDirectoryW" ( _
                         ByVal lpBuffer As Long, _
                         ByVal nSize As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32" _
                         Alias "CreateFileMappingW" ( _
                         ByVal hFile As Long, _
                         ByRef lpFileMappingAttributes As Any, _
                         ByVal flProtect As Long, _
                         ByVal dwMaximumSizeHigh As Long, _
                         ByVal dwMaximumSizeLow As Long, _
                         ByVal lpName As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" ( _
                         ByVal hFileMappingObject As Long, _
                         ByVal dwDesiredAccess As Long, _
                         ByVal dwFileOffsetHigh As Long, _
                         ByVal dwFileOffsetLow As Long, _
                         ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" ( _
                         ByVal lpBaseAddress As Long) As Long
Private Declare Function StrCmpCA Lib "shlwapi" ( _
                         ByRef lpString1 As Any, _
                         ByRef lpString2 As Any) As Long
Private Declare Function RtlGetVersion Lib "ntdll" ( _
                         ByRef pVersion As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" _
                         Alias "GetFileVersionInfoSizeW" ( _
                         ByVal lptstrFilename As Long, _
                         ByRef lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" _
                         Alias "GetFileVersionInfoW" ( _
                         ByVal lptstrFilename As Long, _
                         ByVal dwHandle As Long, _
                         ByVal dwLen As Long, _
                         ByRef lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" _
                         Alias "VerQueryValueW" ( _
                         ByRef pBlock As Any, _
                         ByVal lpSubBlock As Long, _
                         ByRef lpBuffer As Any, _
                         ByRef puLen As Long) As Long
                         
Private Declare Sub memcpy Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)
Private Declare Sub MoveArray Lib "msvbvm60" _
                    Alias "__vbaAryMove" ( _
                    ByRef Destination() As Any, _
                    ByRef Source As Any)
                         
Private m_pCodeBuffer           As Long
Private m_hCurHandle            As Long
Private m_hUser32               As OLE_HANDLE
Private m_pfnZwUserMessageCall  As Long
Private m_bWin8AndAbove         As Boolean

' // Initialize module
Public Function Initialize() As Boolean
    Dim hWow64CPU   As OLE_HANDLE
    Dim lFSRedirect As Long
    Dim lResSize    As Long
    Dim bVersion()  As Byte
    Dim pFixVer     As Long
    Dim lFixLen     As Long
    Dim tFixVer     As VS_FIXEDFILEINFO
    Dim tSysInfo    As SYSTEM_INFO
    
    GetNativeSystemInfo tSysInfo
    
    If tSysInfo.wProcessorArchitecture <> PROCESSOR_ARCHITECTURE_AMD64 Then
        Exit Function
    End If
    
    If m_pCodeBuffer = 0 Then
        
        If Wow64DisableWow64FsRedirection(lFSRedirect) = 0 Then
            Exit Function
        End If
    
        lResSize = GetFileVersionInfoSize(StrPtr("wow64cpu.dll"), 0)
        
        If lResSize > 0 Then
            
            ReDim bVersion(lResSize - 1)
            lResSize = GetFileVersionInfo(StrPtr("wow64cpu.dll"), 0, lResSize, bVersion(0))
            
        End If

        Wow64RevertWow64FsRedirection lFSRedirect
        
        If lResSize = 0 Then
            Exit Function
        End If
        
        If VerQueryValue(bVersion(0), StrPtr("\"), pFixVer, lFixLen) = 0 Then
            Exit Function
        End If

        memcpy tFixVer, ByVal pFixVer, Len(tFixVer)
        
        m_bWin8AndAbove = tFixVer.dwFileVersionMS >= &H60002
        
        m_hCurHandle = OpenProcess(PROCESS_VM_READ, 0, GetCurrentProcessId())
        
        If m_hCurHandle = 0 Then
            Exit Function
        End If
        
        ' // Temporary buffer for caller
        ' // Be careful it doesn't support threading
        ' // To support threading you should ensure atomic access to that buffer
        m_pCodeBuffer = VirtualAlloc(0, 4096, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
        
        If m_pCodeBuffer = 0 Then
            CloseHandle m_hCurHandle
            Exit Function
        End If
        
    End If
    
    Initialize = True
    
End Function

' // Uninitialize module
Public Sub Uninitialize()
    
    If m_hUser32 Then
        UnmapViewOfFile m_hUser32
    End If
    
    If m_hCurHandle Then
        CloseHandle m_hCurHandle
    End If
    
    If m_pCodeBuffer Then
        VirtualFree m_pCodeBuffer, 0, MEM_RELEASE
    End If
    
End Sub

' //
' // Call 64 bit function by pointer
' //
Public Function CallX64( _
                ByVal pfn64 As Currency, _
                ParamArray vArgs() As Variant) As Currency
    Dim bCode()     As Byte             ' // Array to map code
    Dim vArg        As Variant
    Dim lIndex      As Long
    Dim lByteIdx    As Long
    Dim lArgs       As Long
    Dim tArrDesc    As SAFEARRAY1D
    Dim vRet        As Variant
    Dim hr          As Long
    
    If m_pCodeBuffer = 0 Then
        
        ' // Isn't initialized
        Err.Raise 5
        Exit Function
        
    End If
    
    ' // Map array
    tArrDesc.cbElements = 1
    tArrDesc.cDims = 1
    tArrDesc.fFeatures = FADF_AUTO
    tArrDesc.Bounds.cElements = 4096
    tArrDesc.pvData = m_pCodeBuffer
    
    MoveArray bCode(), VarPtr(tArrDesc)
    
    ' // Make x64call
    
    ' // JMP FAR 33:ADDR
    bCode(0) = &HEA
    
    GetMem4 m_pCodeBuffer + 7, bCode(1)
    GetMem2 &H33, bCode(5)
    
    lByteIdx = 7
    
    ' // stack alignment
    
    ' // PUSH RBX
    ' // MOV RBX, SS
    
    ' // --- win7 and below ---
    
    ' // MOV [R12 + 0x1480], R14 ; TlsSlot for stack (R12 - TEB64)
    
    ' // ----------------------
    
    ' // XCHG RSP, R14
    ' // MOV [R13 + STACK_PTR], R14
    ' // PUSH RBP
    ' // MOV RBP, RSP
    ' // AND ESP, 0xFFFFFFF0
    ' // SUB RSP, 0x28 + Args

    If UBound(vArgs) <= 3 Then
        lArgs = 4
    Else
        lArgs = ((UBound(vArgs) - 3) + 1) And &HFFFFFFFE
    End If
    
    lArgs = lArgs * 8 + &H20
    
    PutMem4 bCode(lByteIdx), &HD38C4853:                lByteIdx = lByteIdx + 4
    
    If Not m_bWin8AndAbove Then
        GetMem8 2254060418.0813@, bCode(lByteIdx):      lByteIdx = lByteIdx + 8
    End If
    
    GetMem8 19960132210.0553@, bCode(lByteIdx):         lByteIdx = lByteIdx + 8
    GetMem8 -198047932815474.688@, bCode(lByteIdx):     lByteIdx = lByteIdx + 8
    
    If m_bWin8AndAbove Then
        GetMem4 &H48&, bCode(lByteIdx - 10)
    Else
        GetMem4 &HC8&, bCode(lByteIdx - 10)
    End If
    
    PutMem4 bCode(lByteIdx), &HEC8148F0:                lByteIdx = lByteIdx + 4
    GetMem4 lArgs, bCode(lByteIdx):                     lByteIdx = lByteIdx + 4

    For Each vArg In vArgs
        
        Select Case VarType(vArg)
        Case vbLong, vbString, vbInteger, vbByte, vbBoolean

            Select Case lIndex
            Case 0: GetMem4 &HC1C748, bCode(lByteIdx):  lByteIdx = lByteIdx + 3
            Case 1: GetMem4 &HC2C748, bCode(lByteIdx):  lByteIdx = lByteIdx + 3
            Case 2: GetMem4 &HC0C749, bCode(lByteIdx):  lByteIdx = lByteIdx + 3
            Case 3: GetMem4 &HC1C749, bCode(lByteIdx):  lByteIdx = lByteIdx + 3
            Case Else
            
                GetMem4 &H2444C748, bCode(lByteIdx):    lByteIdx = lByteIdx + 4
                GetMem1 (lIndex - 4) * 8 + &H20, bCode(lByteIdx):   lByteIdx = lByteIdx + 1

            End Select
            
            Select Case VarType(vArg)
            Case vbLong, vbInteger, vbByte, vbBoolean
                GetMem4 CLng(vArg), bCode(lByteIdx):            lByteIdx = lByteIdx + 4
            Case vbString
                GetMem4 ByVal StrPtr(vArg), bCode(lByteIdx):    lByteIdx = lByteIdx + 4
            End Select
            
        Case vbCurrency
        
            Select Case lIndex
            Case 0: GetMem2 &HB948, bCode(lByteIdx):  lByteIdx = lByteIdx + 2
            Case 1: GetMem2 &HBA48, bCode(lByteIdx):  lByteIdx = lByteIdx + 2
            Case 2: GetMem2 &HB849, bCode(lByteIdx):  lByteIdx = lByteIdx + 2
            Case 3: GetMem2 &HB949, bCode(lByteIdx):  lByteIdx = lByteIdx + 2
            Case Else
            
                GetMem2 &HB848, bCode(lByteIdx):      lByteIdx = lByteIdx + 2
                GetMem8 CCur(vArg), bCode(lByteIdx):  lByteIdx = lByteIdx + 8
                GetMem4 &H24448948, bCode(lByteIdx):  lByteIdx = lByteIdx + 4
                GetMem1 (lIndex - 4) * 8 + &H20, bCode(lByteIdx):   lByteIdx = lByteIdx + 1
                
            End Select
            
            If lIndex < 4 Then
                GetMem8 CCur(vArg), bCode(lByteIdx):  lByteIdx = lByteIdx + 8
            End If
        
        Case Else
            
            Err.Raise 13
            Exit Function
            
        End Select
        
        lIndex = lIndex + 1
        
    Next
    
    ' // MOV RAX, pfn: CALL RAX
    GetMem2 &HB848, bCode(lByteIdx):    lByteIdx = lByteIdx + 2
    GetMem8 pfn64, bCode(lByteIdx):     lByteIdx = lByteIdx + 8
    GetMem2 &HD0FF&, bCode(lByteIdx):   lByteIdx = lByteIdx + 2
    
    ' // LEAVE
    ' // XCHG RSP, R14
    ' // MOV SS, RBX
    ' // POP RBX
    GetMem8 661678872152868.7817@, bCode(lByteIdx): lByteIdx = lByteIdx + 8
    
    ' // RAX to EAX/EDX pair
    ' // MOV RDX, RAX
    ' // SHR RDX, 0x20
    GetMem8 926531512503.7384@, bCode(lByteIdx):
    lByteIdx = lByteIdx + 7
    
    ' // JMP FAR 23:
    GetMem2 &H2DFF, bCode(lByteIdx):    lByteIdx = lByteIdx + 2
    GetMem4 0&, bCode(lByteIdx):        lByteIdx = lByteIdx + 4
    GetMem4 m_pCodeBuffer + lByteIdx + 6, bCode(lByteIdx)

    lByteIdx = lByteIdx + 4
    GetMem2 &H23&, bCode(lByteIdx):     lByteIdx = lByteIdx + 2

    bCode(lByteIdx) = &HC3

    hr = DispCallFunc(ByVal 0&, m_pCodeBuffer, 4, vbCurrency, 0, ByVal 0&, ByVal 0&, vRet)

    GetMem4 0&, ByVal ArrPtr(bCode)

    If hr < 0 Then
        Err.Raise hr
        Exit Function
    End If
    
    CallX64 = vRet
    
End Function

' //
' // Get procedure arrdess from 64 bit dll
' //
Public Function GetProcAddress64( _
                ByVal h64Lib As Currency, _
                ByRef sFunctionName As String) As Currency
    Dim lRvaNtHeaders       As Long
    Dim tExportData         As IMAGE_DATA_DIRECTORY
    Dim tExportDirectory    As IMAGE_EXPORT_DIRECTORY
    Dim lIndex              As Long
    Dim p64SymName          As Currency
    Dim tasFunction         As ANSI_STRING64
    Dim tasSymbol           As ANSI_STRING64
    Dim sAnsiString         As String
    Dim lOrdinal            As Long
    Dim p64Address          As Currency
    
    If h64Lib = 0 Then
        
        h64Lib = GetModuleHandle64(vbNullString)
            
        If h64Lib = 0 Then
        
            Err.Raise 5
            Exit Function
            
        End If
            
    End If

    sAnsiString = StrConv(sFunctionName, vbFromUnicode)
    
    GetMem4 StrPtr(sAnsiString), tasFunction.lpBuffer
    tasFunction.Length = LenB(sAnsiString)
    tasFunction.MaxLength = tasFunction.Length + 1
    
    ReadMem64 VarPtr(lRvaNtHeaders), h64Lib + 0.006@, Len(lRvaNtHeaders)
    ReadMem64 VarPtr(tExportData), h64Lib + lRvaNtHeaders / 10000 + 0.0136@, Len(tExportData)
    
    If tExportData.VirtualAddress = 0 Or tExportData.Size = 0 Then
        Err.Raise 453
        Exit Function
    End If
    
    ReadMem64 VarPtr(tExportDirectory), h64Lib + tExportData.VirtualAddress / 10000, Len(tExportDirectory)
    
    For lIndex = 0 To tExportDirectory.NumberOfNames - 1
        
        p64SymName = 0
        
        ReadMem64 VarPtr(p64SymName), (tExportDirectory.AddressOfNames + lIndex * 4) / 10000 + h64Lib, 4
        
        p64SymName = p64SymName + h64Lib
        
        tasSymbol.Length = StringLen64(p64SymName) * 10000
        tasSymbol.MaxLength = tasSymbol.Length
        tasSymbol.lpBuffer = p64SymName
        
        If CompareAnsiStrings64(tasFunction, tasSymbol, True) = 0 Then
            
            ReadMem64 VarPtr(lOrdinal), (tExportDirectory.AddressOfNameOrdinals + lIndex * 2) / 10000 + h64Lib, 2
            ReadMem64 VarPtr(p64Address), (tExportDirectory.AddressOfFunctions + lOrdinal * 4) / 10000 + h64Lib, 4
            
            GetProcAddress64 = p64Address + h64Lib
            
            Exit For
            
        End If
        
    Next

End Function

' //
' // Get 64-bit lib handle
' //
Public Property Get GetModuleHandle64( _
                    ByRef sLib As String) As Currency
    Dim tPBI64          As PROCESS_BASIC_INFORMATION64
    Dim lStatus         As Long
    Dim p64LdrData      As Currency
    Dim p64ListEntry    As Currency
    Dim p64LdrEntry     As Currency
    Dim p64DllName      As Currency
    Dim tusDll          As UNICODE_STRING64
    Dim tusLib          As UNICODE_STRING64

    GetMem4 StrPtr(sLib), tusDll.lpBuffer ' // Address
    tusDll.Length = LenB(sLib)
    tusDll.MaxLength = tusDll.Length + 2
    
    ' // We need 64-bit PEB
    lStatus = NtWow64QueryInformationProcess64(-1, ProcessBasicInformation, tPBI64, Len(tPBI64), 0)
    
    If lStatus < 0 Then
        Err.Raise lStatus
        Exit Property
    End If
    
    ' // Read PEB.Ldr
    ReadMem64 VarPtr(p64LdrData), tPBI64.PebBaseAddress + 0.0024@, Len(p64LdrData)
    
    p64ListEntry = p64LdrData + 0.0016@ ' // PEB_LDR_DATA.InLoadOrderModuleList.Flink
    
    ' // *PEB_LDR_DATA.InLoadOrderModuleList.Flink
    ReadMem64 VarPtr(p64LdrEntry), p64ListEntry, Len(p64LdrEntry)

    Do
        
        p64DllName = p64LdrEntry + 0.0088@ ' // LDR_DATA_TABLE_ENTRY.BaseDllName
        
        If Len(sLib) = 0 Then
            
            ReadMem64 VarPtr(GetModuleHandle64), p64LdrEntry + 0.0048@, Len(GetModuleHandle64)
            Exit Do
            
        Else
            
            ReadMem64 VarPtr(tusLib), p64DllName, Len(tusLib)
            
            If CompareUnicodeStrings64(tusLib, tusDll) = 0 Then
                
                ReadMem64 VarPtr(GetModuleHandle64), p64LdrEntry + 0.0048@, Len(GetModuleHandle64)
                Exit Do
                
            End If
        
        End If
        
        ReadMem64 VarPtr(p64LdrEntry), p64LdrEntry, Len(p64LdrEntry)

    Loop Until p64ListEntry = p64LdrEntry
    
End Property

' // Read memory at specified 64-bit address
Public Sub ReadMem64( _
           ByVal pTo As Long, _
           ByVal p64From As Currency, _
           ByVal lSize As Long)
    Dim lStatus As Long
    
    lStatus = NtWow64ReadVirtualMemory64(m_hCurHandle, p64From, ByVal pTo, lSize / 10000, 0)

    If lStatus < 0 Then
        Err.Raise lStatus
        Exit Sub
    End If
 
End Sub

' // Send message and return 64 bit result
Public Function SendMessageW64( _
                ByVal hwnd As OLE_HANDLE, _
                ByVal lMsg As Long, _
                ByVal p64wParam As Currency, _
                ByVal p64lParam As Currency) As Currency
    Dim hr          As Long
    Dim lFID        As Long
    Dim p64Result   As Currency
    Dim p64Fn       As Currency
    
    If m_pfnZwUserMessageCall = 0 Then
    
        hr = GetZwUserMessageCallAddress(m_pfnZwUserMessageCall, m_hUser32)
        
        If hr < 0 Then
            Err.Raise hr
        End If
        
    End If
    
    PutMem4 p64Fn, m_pfnZwUserMessageCall

    SendMessageW64 = CallX64(p64Fn, hwnd, lMsg, p64wParam, p64lParam, 0&, &H2B1, 0&)
    
End Function

Private Function GetZwUserMessageCallAddress( _
                 ByRef pfnRet As Long, _
                 ByRef hLib As OLE_HANDLE) As Long
    Dim hModule As OLE_HANDLE
    Dim hr      As Long
    Dim pfn     As Long
    
    hr = MapModule("win32u.dll", hModule)
    
    If hr < 0 Then
            
        hr = MapModule("user32.dll", hModule)
        
        If hr < 0 Then
            GetZwUserMessageCallAddress = hr
            Exit Function
        End If
        
        pfn = SearchExportInLib64(hModule, "gapfnScSendMessage")
        
        If pfn = 0 Then
            hr = &H8007007F
            GoTo CleanUp
        End If
        
        pfn = ZwUserMessageCallAddressFromgapfnScSendMessage(hModule, pfn)
        
    Else
        pfn = SearchExportInLib64(hModule, "NtUserMessageCall")
    End If

    If pfn = 0 Then
        hr = &H8007007F
        GoTo CleanUp
    End If
         
    pfnRet = pfn
    hLib = hModule
    
CleanUp:
    
    If hr < 0 Then
        
        If hModule Then
            UnmapViewOfFile hModule
        End If

    End If
        
    GetZwUserMessageCallAddress = hr
        
End Function

Private Function ZwUserMessageCallAddressFromgapfnScSendMessage( _
                 ByVal hUser32 As Long, _
                 ByVal pfn As Long) As Long
    Dim p64Offset       As Currency
    Dim p64Base         As Currency
    Dim lRvaNtHeaders   As Long
    Dim lRVAFunc        As Long
    Dim tFileHdr        As IMAGE_FILE_HEADER
    Dim tSections()     As IMAGE_SECTION_HEADER
    Dim pAddress        As Long
    
    GetMem4 ByVal hUser32 + &H3C, lRvaNtHeaders
    GetMem8 ByVal lRvaNtHeaders + hUser32 + &H30, p64Base
    GetMem8 ByVal pfn, p64Offset
    GetMem4 p64Offset - p64Base, lRVAFunc
    
    memcpy tFileHdr, ByVal hUser32 + lRvaNtHeaders + 4, Len(tFileHdr)
    
    ReDim tSections(tFileHdr.NumberOfSections - 1)
    
    memcpy tSections(0), ByVal hUser32 + lRvaNtHeaders + &H108, Len(tSections(0)) * tFileHdr.NumberOfSections
    
    ZwUserMessageCallAddressFromgapfnScSendMessage = Rva2Raw(tSections(), lRVAFunc) + hUser32
    
End Function

Private Function SearchExportInLib64( _
                 ByVal hLib As OLE_HANDLE, _
                 ByRef sName As String) As Long
    Dim lRvaNtHeaders       As Long
    Dim tFileHdr            As IMAGE_FILE_HEADER
    Dim tExportDir          As IMAGE_DATA_DIRECTORY
    Dim tSections()         As IMAGE_SECTION_HEADER
    Dim tExportDirectory    As IMAGE_EXPORT_DIRECTORY
    Dim lI                  As Long
    Dim lJ                  As Long
    Dim pNames              As Long
    Dim pFunctionName       As Long
    Dim sNameANSI           As String
    Dim lOrdinal            As Long
    Dim lFnRVA              As Long
    
    sNameANSI = StrConv(sName, vbFromUnicode)
    
    GetMem4 ByVal hLib + &H3C, lRvaNtHeaders
    
    memcpy tFileHdr, ByVal hLib + lRvaNtHeaders + 4, Len(tFileHdr)
    memcpy tExportDir, ByVal hLib + lRvaNtHeaders + 136, Len(tExportDir)
    
    If tExportDir.Size <= 0 Or tExportDir.VirtualAddress <= 0 Then
        Exit Function
    End If
    
    ReDim tSections(tFileHdr.NumberOfSections - 1)
    
    memcpy tSections(0), ByVal hLib + lRvaNtHeaders + &H108, Len(tSections(0)) * tFileHdr.NumberOfSections
    memcpy tExportDirectory, ByVal hLib + Rva2Raw(tSections(), tExportDir.VirtualAddress), Len(tExportDirectory)
        
    pNames = Rva2Raw(tSections(), tExportDirectory.AddressOfNames) + hLib
        
    Do
        
        lI = tExportDirectory.NumberOfNames \ 2
            
        GetMem4 ByVal pNames + (lI + lJ) * 4, pFunctionName
        pFunctionName = Rva2Raw(tSections(), pFunctionName) + hLib
        
        Select Case StrCmpCA(ByVal StrPtr(sNameANSI), ByVal pFunctionName)
        Case 0
            
            GetMem2 ByVal Rva2Raw(tSections(), tExportDirectory.AddressOfNameOrdinals) + hLib + (lI + lJ) * 2, lOrdinal
            GetMem4 ByVal Rva2Raw(tSections(), tExportDirectory.AddressOfFunctions) + hLib + lOrdinal * 4, lFnRVA
            
            If lFnRVA < tExportDir.VirtualAddress Or lFnRVA >= tExportDir.VirtualAddress + tExportDir.Size Then
                SearchExportInLib64 = Rva2Raw(tSections(), lFnRVA) + hLib
            End If
            
            Exit Function
            
        Case Is > 0
            tExportDirectory.NumberOfNames = tExportDirectory.NumberOfNames - lI
            lJ = lJ + lI
        Case Else
            tExportDirectory.NumberOfNames = tExportDirectory.NumberOfNames - lI
        End Select
        
    Loop While lI

End Function

Private Function Rva2Raw( _
                 ByRef tSections() As IMAGE_SECTION_HEADER, _
                 ByVal lRVA As Long) As Long
    Dim lIndex  As Long
    
    For lIndex = 0 To UBound(tSections)
        If lRVA >= tSections(lIndex).VirtualAddress And lRVA < tSections(lIndex).VirtualAddress + tSections(lIndex).VirtualSize Then
            Rva2Raw = tSections(lIndex).PointerToRawData + (lRVA - tSections(lIndex).VirtualAddress)
            Exit Function
        End If
    Next
    
    Rva2Raw = lRVA
    
End Function

Private Function MapModule( _
                 ByRef sName As String, _
                 ByRef hResult As OLE_HANDLE) As Long
    Dim hMap            As OLE_HANDLE
    Dim hFile           As OLE_HANDLE
    Dim sSysPath        As String
    Dim lSize           As Long
    Dim pAddress        As OLE_HANDLE
    Dim hr              As Long
    Dim lRvaNtHeaders   As Long
    Dim lMachine        As Long
    Dim lFSRedirect     As Long
    
    sSysPath = Space$(MAX_PATH)
    lSize = GetSystemDirectory(StrPtr(sSysPath), Len(sSysPath) + 1)
    sSysPath = Left$(sSysPath, lSize)
    
    If Wow64DisableWow64FsRedirection(lFSRedirect) = 0 Then
        MapModule = &H80070000 Or (Err.LastDllError And &HFFFF&)
        Exit Function
    End If
    
    hFile = CreateFile(StrPtr(sSysPath & "\" & sName), GENERIC_READ Or GENERIC_EXECUTE, 5, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    Wow64RevertWow64FsRedirection lFSRedirect
    
    If hFile = INVALID_HANDLE_VALUE Then
        MapModule = &H80070000 Or (Err.LastDllError And &HFFFF&)
        Exit Function
    End If
    
    hMap = CreateFileMapping(hFile, ByVal 0&, PAGE_EXECUTE_READ, 0, 0, 0)
    
    If hMap = 0 Then
        hr = &H80070000 Or (Err.LastDllError And &HFFFF&)
        GoTo CleanUp
    End If
    
    pAddress = MapViewOfFile(hMap, FILE_MAP_READ Or FILE_MAP_EXECUTE, 0, 0, 0)
    
    If pAddress = 0 Then
        hr = &H80070000 Or (Err.LastDllError And &HFFFF&)
        GoTo CleanUp
    End If
    
    GetMem4 ByVal pAddress + &H3C, lRvaNtHeaders
    GetMem2 ByVal pAddress + lRvaNtHeaders + 4, lMachine
    
    If lMachine <> &H8664& Then
        hr = &H8000FFFF
        GoTo CleanUp
    End If
    
CleanUp:
    
    If hMap Then
        CloseHandle hMap
    End If
    
    If hFile Then
        CloseHandle hFile
    End If
    
    If hr < 0 Then
        If pAddress Then
            UnmapViewOfFile pAddress
        End If
    Else
        hResult = pAddress
    End If
    
    MapModule = hr
    
End Function

' // Get SendMessage address
' // Get null-terminated string length
Private Function StringLen64( _
                 ByVal p64 As Currency) As Currency
    Dim pAddrPair(1)    As Long
    Dim bPage()         As Byte
    Dim lSize           As Long
    Dim lStatus         As Long
    Dim lIndex          As Long
    Dim p64Start        As Currency
    
    p64Start = p64
    
    GetMem8 p64, pAddrPair(0)
    
    ' // Get number of bytes to end page boundry
    lSize = &H1000 - (pAddrPair(0) And &HFFF)
    
    Do
        
        ' // Read page
        ReDim Preserve bPage(lSize - 1)
        
        lStatus = NtWow64ReadVirtualMemory64(m_hCurHandle, p64, bPage(0), lSize / 10000, 0)
        
        If lStatus < 0 Then
            Err.Raise lStatus
            Exit Function
        End If
    
        For lIndex = 0 To lSize - 1
            
            ' // Test for null terminal
            If bPage(lIndex) = 0 Then
                
                StringLen64 = (p64 + lIndex / 10000) - p64Start
                Exit Do
                
            End If

        Next
        
        ' // Next page
        p64 = p64 + lSize / 10000
        
        lSize = 4096
        
    Loop While True
                 
End Function

' // Compare 2 ANSI strings
Private Function CompareAnsiStrings64( _
                 ByRef tasStr1 As ANSI_STRING64, _
                 ByRef tasStr2 As ANSI_STRING64, _
                 Optional ByVal bCaseSensitive As Boolean) As Long
    Dim bBuf1() As Byte
    Dim bBuf2() As Byte

    If tasStr1.Length > 0 Then
        
        ReDim bBuf1(tasStr1.Length)
        ReadMem64 VarPtr(bBuf1(0)), tasStr1.lpBuffer, tasStr1.Length
        
    End If
    
    If tasStr2.Length > 0 Then
    
        ReDim bBuf2(tasStr2.Length)
        ReadMem64 VarPtr(bBuf2(0)), tasStr2.lpBuffer, tasStr2.Length
        
    End If
    
    If bCaseSensitive Then
        CompareAnsiStrings64 = lstrcmp(bBuf1(0), bBuf2(0))
    Else
        CompareAnsiStrings64 = lstrcmpi(bBuf1(0), bBuf2(0))
    End If
    
End Function

' // Compare 2 strings
Private Function CompareUnicodeStrings64( _
                 ByRef tusStr1 As UNICODE_STRING64, _
                 ByRef tusStr2 As UNICODE_STRING64, _
                 Optional ByVal bCaseSensitive As Boolean) As Long
    Dim bBuf1() As Byte
    Dim bBuf2() As Byte

    If tusStr1.Length > 0 Then
        
        ReDim bBuf1(tusStr1.Length - 1)
        ReadMem64 VarPtr(bBuf1(0)), tusStr1.lpBuffer, tusStr1.Length
        
    End If
    
    If tusStr2.Length > 0 Then
    
        ReDim bBuf2(tusStr2.Length - 1)
        ReadMem64 VarPtr(bBuf2(0)), tusStr2.lpBuffer, tusStr2.Length
        
    End If

    If bCaseSensitive Then
        CompareUnicodeStrings64 = StrComp(bBuf1, bBuf2, vbBinaryCompare)
    Else
        CompareUnicodeStrings64 = StrComp(bBuf1, bBuf2, vbTextCompare)
    End If
    
End Function

