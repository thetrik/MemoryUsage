VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Memory usage by The trick"
   ClientHeight    =   10320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwInfo 
      Height          =   2190
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   3863
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Work (Kb)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Page file (Kb)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Page fault"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Peak page file (Kb)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Peak working (Kb)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Non-paged pool (Kb)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Paged pool (Kb)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Peak non-paged pool (Kb)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Peak paged pool (Kb)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer tmrTimer 
      Interval        =   1000
      Left            =   390
      Top             =   5340
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // frmMain.frm
' // Show information about memory usage
' // by The trick 2014 - 2020
' //

Option Explicit

Private Const MAX_PATH          As Long = 260
Private Const ProcessVmCounters As Long = 3

Private Type VM_COUNTERS64
    PeakVirtualSize             As Currency
    VirtualSize                 As Currency
    PageFaultCount              As Long
    lPad1                       As Long
    PeakWorkingSetSize          As Currency
    WorkingSetSize              As Currency
    QuotaPeakPagedPoolUsage     As Currency
    QuotaPagedPoolUsage         As Currency
    QuotaPeakNonPagedPoolUsage  As Currency
    QuotaNonPagedPoolUsage      As Currency
    PagefileUsage               As Currency
    PeakPagefileUsage           As Currency
End Type

Private Type PROCESS_MEMORY_COUNTERS64
    cb                          As Long
    PageFaultCount              As Long
    PeakWorkingSetSize          As Currency
    WorkingSetSize              As Currency
    QuotaPeakPagedPoolUsage     As Currency
    QuotaPagedPoolUsage         As Currency
    QuotaPeakNonPagedPoolUsage  As Currency
    QuotaNonPagedPoolUsage      As Currency
    PagefileUsage               As Currency
    PeakPagefileUsage           As Currency
End Type

Private Type PROCESS_MEMORY_COUNTERS
    cb                          As Long
    PageFaultCount              As Long
    PeakWorkingSetSize          As Long
    WorkingSetSize              As Long
    QuotaPeakPagedPoolUsage     As Long
    QuotaPagedPoolUsage         As Long
    QuotaPeakNonPagedPoolUsage  As Long
    QuotaNonPagedPoolUsage      As Long
    PagefileUsage               As Long
    PeakPagefileUsage           As Long
End Type

Private Type PROCESSENTRY32
    dwSize                      As Long
    cntUsage                    As Long
    th32ProcessID               As Long
    th32DefaultHeapID           As Long
    th32ModuleID                As Long
    cntThreads                  As Long
    th32ParentProcessID         As Long
    pcPriClassBase              As Long
    dwFlags                     As Long
    szExeFile                   As String * MAX_PATH
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize         As Long
    dwMajorVersion              As Long
    dwMinorVersion              As Long
    dwBuildNumber               As Long
    dwPlatformId                As Long
    szCSDVersion                As String * 128
End Type

Private Type tListEntry
    lPID                        As Long
    sExeName                    As String
    tCounters                   As PROCESS_MEMORY_COUNTERS64
End Type

Private Declare Function GetVersionEx Lib "kernel32" _
                         Alias "GetVersionExA" ( _
                         ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function Process32First Lib "kernel32" ( _
                         ByVal hSnapshot As Long, _
                         ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" ( _
                         ByVal hSnapshot As Long, _
                         ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "kernel32" ( _
                         ByVal dwDesiredAccess As Long, _
                         ByVal bInheritHandle As Long, _
                         ByVal dwProcessId As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
                         ByVal dwFlags As Long, _
                         ByVal th32ProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib "psapi.dll" ( _
                         ByVal lHandle As Long, _
                         ByRef lpStructure As PROCESS_MEMORY_COUNTERS, _
                         ByVal lSize As Long) As Long
Private Declare Function IsWow64Process Lib "kernel32.dll" ( _
                         ByVal lHandle As Long, _
                         ByRef Wow64Process As Long) As Long
Private Declare Function StrFormatKBSize Lib "shlwapi" _
                         Alias "StrFormatKBSizeW" ( _
                         ByVal qdw As Currency, _
                         ByVal pszBuf As Long, _
                         ByVal cchBuf As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60.dll" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Sub ZeroMemory Lib "kernel32" _
                    Alias "RtlZeroMemory" ( _
                    ByRef Destination As Any, _
                    ByVal Length As Long)
                         
Private Const TH32CS_SNAPPROCESS                  As Long = 2
Private Const PROCESS_QUERY_LIMITED_INFORMATION   As Long = &H1000
Private Const PROCESS_QUERY_INFORMATION           As Long = &H400
Private Const INVALID_HANDLE_VALUE                As Long = -1

Private m_bIsVistaAndLater              As Boolean
Private m_bIs64BitEnvironment           As Boolean
Private m_p64NtQueryInformationProcess  As Currency
Private m_lSortKey                      As Long
Private m_eSortOrder                    As ListSortOrderConstants
Private m_tProcessList()                As tListEntry
Private m_lProcessCount                 As Long
Private m_lSelPID                       As Long

Private Sub Form_Load()
    Dim tVer As OSVERSIONINFO

    tVer.dwOSVersionInfoSize = Len(tVer)
    
    GetVersionEx tVer
    m_bIsVistaAndLater = tVer.dwMajorVersion >= 6
    m_bIs64BitEnvironment = Is64BitEnv

    If m_bIs64BitEnvironment Then
        If Not modX64Call.Initialize() Then
            MsgBox "Unable to initialize x64caller", vbCritical
        End If
    End If
    
    m_lSortKey = 0

    Call tmrTimer_Timer
    
End Sub

Private Sub Form_Resize()
    If Me.ScaleWidth > 200 And Me.ScaleHeight > 200 Then
        lvwInfo.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
    End If
End Sub

Private Function Is64BitEnv() As Boolean
    Dim lIsWow64    As Long
    
    On Error GoTo api_not_found
    
    If IsWow64Process(-1, lIsWow64) = 0 Then
        lIsWow64 = 0
    End If
    
    Is64BitEnv = lIsWow64
    
api_not_found:
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    modX64Call.Uninitialize
End Sub

Private Sub lvwInfo_ColumnClick( _
            ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    If m_lSortKey = ColumnHeader.Index - 1 Then
        m_eSortOrder = (m_eSortOrder + 1) And 1
    Else
        m_eSortOrder = lvwAscending
    End If
    
    m_lSortKey = ColumnHeader.Index - 1
    
    qSort 0, m_lProcessCount - 1
    
    ' // Update
    tmrTimer_Timer
    
End Sub

Private Sub lvwInfo_ItemClick( _
            ByVal Item As MSComctlLib.ListItem)
    m_lSelPID = Item.Tag
End Sub

Private Sub tmrTimer_Timer()
    Dim hSnap       As Long
    Dim hProcess    As Long
    Dim tPEEntry    As PROCESSENTRY32
    Dim lIndex      As Long
    Dim cLstItem    As ListItem
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If hSnap = INVALID_HANDLE_VALUE Then Exit Sub
    
    ReDim m_tProcessList(99)
    
    m_lProcessCount = 0
    
    tPEEntry.dwSize = Len(tPEEntry)
    
    If Process32First(hSnap, tPEEntry) Then
    
        Do
            
            hProcess = OpenProcess(IIf(m_bIsVistaAndLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), _
                                False, tPEEntry.th32ProcessID)
            If hProcess Then
            
                m_tProcessList(lIndex).tCounters.cb = Len(m_tProcessList(lIndex).tCounters)
                
                GetProcessMemoryInfo64 hProcess, m_tProcessList(lIndex).tCounters
                
                m_tProcessList(lIndex).lPID = tPEEntry.th32ProcessID
                m_tProcessList(lIndex).sExeName = Left$(tPEEntry.szExeFile, InStr(1, tPEEntry.szExeFile, vbNullChar))

                CloseHandle hProcess
                
                lIndex = lIndex + 1
                
                If lIndex > UBound(m_tProcessList) Then
                    ReDim Preserve m_tProcessList(lIndex + 100)
                End If
            
            End If
            
        Loop While Process32Next(hSnap, tPEEntry)
        
    End If
    
    CloseHandle hSnap
    
    m_lProcessCount = lIndex
    
    qSort 0, m_lProcessCount - 1
    
    For lIndex = 0 To m_lProcessCount - 1

        If lIndex >= lvwInfo.ListItems.Count Then
            Set cLstItem = lvwInfo.ListItems.Add
        Else
            Set cLstItem = lvwInfo.ListItems(lIndex + 1)
        End If

        With cLstItem
            
            .Text = m_tProcessList(lIndex).sExeName
            .SubItems(1) = m_tProcessList(lIndex).lPID & " (0x" & Hex$(m_tProcessList(lIndex).lPID) & ")"
            .SubItems(2) = FormatSize(m_tProcessList(lIndex).tCounters.WorkingSetSize)
            .SubItems(3) = FormatSize(m_tProcessList(lIndex).tCounters.PagefileUsage)
            .SubItems(4) = m_tProcessList(lIndex).tCounters.PageFaultCount
            .SubItems(5) = FormatSize(m_tProcessList(lIndex).tCounters.PeakPagefileUsage)
            .SubItems(6) = FormatSize(m_tProcessList(lIndex).tCounters.PeakWorkingSetSize)
            .SubItems(7) = FormatSize(m_tProcessList(lIndex).tCounters.QuotaNonPagedPoolUsage)
            .SubItems(8) = FormatSize(m_tProcessList(lIndex).tCounters.QuotaPagedPoolUsage)
            .SubItems(9) = FormatSize(m_tProcessList(lIndex).tCounters.QuotaPeakNonPagedPoolUsage)
            .SubItems(10) = FormatSize(m_tProcessList(lIndex).tCounters.QuotaPeakPagedPoolUsage)
            
            .Tag = m_tProcessList(lIndex).lPID
            
            If m_lSelPID = m_tProcessList(lIndex).lPID Then
                .Selected = True
            End If
            
        End With

    Next
    
    If m_lProcessCount < lvwInfo.ListItems.Count Then
        Do Until lvwInfo.ListItems.Count = m_lProcessCount
            lvwInfo.ListItems.Remove lvwInfo.ListItems.Count
        Loop
    End If
    
End Sub

Private Sub qSort( _
            ByVal lLow As Long, _
            ByVal lHigh As Long)
    Dim lI As Long, lJ As Long, lM As Long, tS As tListEntry
    
    lI = lLow: lJ = lHigh: lM = (lI + lJ) \ 2
    
    Do Until lI > lJ
    
        Do While qSortCmpFn(m_tProcessList(lI), m_tProcessList(lM)) = -1
            lI = lI + 1
        Loop
        
        Do While qSortCmpFn(m_tProcessList(lJ), m_tProcessList(lM)) = 1
            lJ = lJ - 1
        Loop
        
        If lI <= lJ Then
        
            If lI = lM Then
                lM = lJ
            ElseIf lJ = lM Then
                lM = lI
            End If
            
            tS = m_tProcessList(lJ)
            m_tProcessList(lJ) = m_tProcessList(lI)
            m_tProcessList(lI) = tS
            
            lI = lI + 1: lJ = lJ - 1
            
        End If
        
    Loop
    
    If lLow < lJ Then qSort lLow, lJ
    
    If lI < lHigh Then qSort lI, lHigh
    
End Sub

Private Function qSortCmpFn( _
                 ByRef t1 As tListEntry, _
                 ByRef t2 As tListEntry) As Long
    
    Select Case m_lSortKey
    Case 0: qSortCmpFn = StrComp(t1.sExeName, t2.sExeName, vbTextCompare)
    Case 1: qSortCmpFn = Sgn(t1.lPID - t2.lPID)
    Case 2: qSortCmpFn = Sgn(t1.tCounters.WorkingSetSize - t2.tCounters.WorkingSetSize)
    Case 3: qSortCmpFn = Sgn(t1.tCounters.PagefileUsage - t2.tCounters.PagefileUsage)
    Case 4: qSortCmpFn = Sgn(t1.tCounters.PageFaultCount - t2.tCounters.PageFaultCount)
    Case 5: qSortCmpFn = Sgn(t1.tCounters.PeakPagefileUsage - t2.tCounters.PeakPagefileUsage)
    Case 6: qSortCmpFn = Sgn(t1.tCounters.PeakWorkingSetSize - t2.tCounters.PeakWorkingSetSize)
    Case 7: qSortCmpFn = Sgn(t1.tCounters.QuotaNonPagedPoolUsage - t2.tCounters.QuotaNonPagedPoolUsage)
    Case 8: qSortCmpFn = Sgn(t1.tCounters.QuotaPagedPoolUsage - t2.tCounters.QuotaPagedPoolUsage)
    Case 9: qSortCmpFn = Sgn(t1.tCounters.QuotaPeakNonPagedPoolUsage - t2.tCounters.QuotaPeakNonPagedPoolUsage)
    Case 10: qSortCmpFn = Sgn(t1.tCounters.QuotaPeakPagedPoolUsage - t2.tCounters.QuotaPeakPagedPoolUsage)
    
    End Select
    
    If m_eSortOrder = lvwDescending Then
        qSortCmpFn = -qSortCmpFn
    End If
    
End Function

Private Function FormatSize( _
                 ByVal cValue As Currency) As String
    
    FormatSize = Space$(32)
    
    If StrFormatKBSize(cValue, StrPtr(FormatSize), Len(FormatSize)) Then
        FormatSize = Left$(FormatSize, InStr(1, FormatSize, vbNullChar) - 1)
    Else
        FormatSize = "ERROR"
    End If
    
End Function

Private Function GetProcessMemoryInfo64( _
                 ByVal hProcess As Long, _
                 ByRef tMemInfo As PROCESS_MEMORY_COUNTERS64) As Long
    Dim tMemInfo32  As PROCESS_MEMORY_COUNTERS
    Dim tVmCounters As VM_COUNTERS64
    Dim lRetLen     As Long
    Dim lStatus     As Long
    
    If m_bIs64BitEnvironment Then
            
        On Error GoTo error_handler
            
        If m_p64NtQueryInformationProcess = 0 Then
            
            m_p64NtQueryInformationProcess = GetProcAddress64(GetModuleHandle64("ntdll.dll"), "NtQueryInformationProcess")
            
            If m_p64NtQueryInformationProcess = 0 Then Exit Function
            
        End If
        
        GetMem4 CallX64(m_p64NtQueryInformationProcess, hProcess, ProcessVmCounters, _
                VarPtr(tVmCounters), Len(tVmCounters), VarPtr(lRetLen)), lStatus
        
        If lStatus < 0 Then Exit Function
        
        With tMemInfo
        
            .cb = Len(tMemInfo)
            .PageFaultCount = tVmCounters.PageFaultCount
            .PagefileUsage = tVmCounters.PagefileUsage
            .PeakPagefileUsage = tVmCounters.PeakPagefileUsage
            .PeakWorkingSetSize = tVmCounters.PeakWorkingSetSize
            .QuotaNonPagedPoolUsage = tVmCounters.QuotaNonPagedPoolUsage
            .QuotaPagedPoolUsage = tVmCounters.QuotaPagedPoolUsage
            .QuotaPeakNonPagedPoolUsage = tVmCounters.QuotaPeakNonPagedPoolUsage
            .QuotaPeakPagedPoolUsage = tVmCounters.QuotaPeakPagedPoolUsage
            .WorkingSetSize = tVmCounters.WorkingSetSize
        
        End With
        
        GetProcessMemoryInfo64 = 1
        
    Else
        
        tMemInfo32.cb = Len(tMemInfo32)
        
        GetProcessMemoryInfo64 = GetProcessMemoryInfo(hProcess, tMemInfo32, Len(tMemInfo32))
        
        ZeroMemory tMemInfo, Len(tMemInfo)
        
        With tMemInfo
        
            .cb = Len(tMemInfo)
            .PageFaultCount = tMemInfo32.PageFaultCount
            GetMem4 tMemInfo32.PagefileUsage, .PagefileUsage
            GetMem4 tMemInfo32.PeakPagefileUsage, .PeakPagefileUsage
            GetMem4 tMemInfo32.PeakWorkingSetSize, .PeakWorkingSetSize
            GetMem4 tMemInfo32.QuotaNonPagedPoolUsage, .QuotaNonPagedPoolUsage
            GetMem4 tMemInfo32.QuotaPagedPoolUsage, .QuotaPagedPoolUsage
            GetMem4 tMemInfo32.QuotaPeakNonPagedPoolUsage, .QuotaPeakNonPagedPoolUsage
            GetMem4 tMemInfo32.QuotaPeakPagedPoolUsage, .QuotaPeakPagedPoolUsage
            GetMem4 tMemInfo32.WorkingSetSize, .WorkingSetSize
        
        End With
        
    End If
    
error_handler:
    
End Function

