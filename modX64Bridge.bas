Attribute VB_Name = "modX64Bridge"
' modX64Bridge.bas
' Thu goi X64Bridge.dll tu VB6 (32-bit)
' Ghi chu khong dau vi VBE bi loi font Unicode
'
' Yeu cau:
' - VB6 IDE hoac compiled EXE (32-bit)
' - X64Bridge.dll (Win32) dat cung thu muc EXE hoac trong PATH
' - Windows 64-bit (WOW64)
Option Explicit
' ============================================================
' Kieu du lieu
' ============================================================
Private Type X64Param
    lo As Long
    hi As Long
End Type
' ============================================================
' Khai bao API — VB6 dung Declare khong co PtrSafe
' VB6 String truyen ByVal = PAnsiChar (ANSI, khac VBA Unicode)
' ============================================================
' Vong doi
Private Declare Function X64Init Lib "X64Bridge.dll" () As Long
Private Declare Sub X64Free Lib "X64Bridge.dll" ()
' Goi ham 64-bit (ANSI — phu hop voi VB6 String)
' API_Invoke duoc export voi alias "X64Invoke"
Private Declare Function X64Invoke Lib "X64Bridge.dll" (ByVal sDll As String, ByVal sFuncName As String, ByVal nArgs As Long, ByRef pParams As X64Param, ByRef rLo As Long, ByRef rHi As Long) As Long
' Buffer (cap phat vung nho de truyen pointer vao ham 64-bit)
Private Declare Function X64BufAllocLo Lib "X64Bridge.dll" (ByVal nBytes As Long) As Long
Private Declare Function X64BufAllocHi Lib "X64Bridge.dll" (ByVal nBytes As Long) As Long
Private Declare Sub X64BufFree Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long)
Private Declare Sub X64BufWriteAnsi Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long, ByVal s As String)
Private Declare Sub X64BufWriteUnicode Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long, ByVal pW As Long)
Private Declare Function X64BufReadDW Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long, ByVal ofs As Long) As Long
' Debug
Private Declare Function X64Version Lib "X64Bridge.dll" () As Long
Private Declare Function X64DebugNtdll Lib "X64Bridge.dll" (ByRef modLo As Long, ByRef modHi As Long, ByRef pfnLo As Long, ByRef pfnHi As Long) As Long
Private Declare Function X64DebugInitState Lib "X64Bridge.dll" (ByRef hProc As Long, ByRef pCode As Long) As Long
Rem SAI � VB6 truyen ByVal String = PAnsiChar, khong phai PWideChar
Rem sai Private Declare Function X64DebugCache Lib "X64Bridge.dll" (ByVal sDll As String, ByRef pLo As Long, ByRef pHi As Long) As Long
Rem DUNG � phai truyen Long (dia chi) vi Delphi nhan PWideChar
Public Declare Function X64DebugCache Lib "X64Bridge.dll" (ByVal sDll As Long, ByRef pLo As Long, ByRef pHi As Long) As Long
'
Private Declare Function X64DebugReadMem Lib "X64Bridge.dll" (ByRef ntstatus As Long) As Long
' Ma loi
Private Const X64_OK          As Long = 0
Private Const X64_ERR_NOINIT  As Long = 1
Private Const X64_ERR_NODLL   As Long = 2
Private Const X64_ERR_NOPROC  As Long = 3
Private Const X64_ERR_TOOMANY As Long = 4
Private Const X64_ERR_EXCEPT  As Long = 5
' ============================================================
' Helper
' ============================================================
Private Function Hex64(lo As Long, hi As Long) As String
    If hi = 0 Then
        Hex64 = "0x" & Right$("00000000" & Hex$(lo), 8)
    Else
        Hex64 = "0x" & Right$("00000000" & Hex$(hi), 8) & Right$("00000000" & Hex$(lo), 8)
    End If
End Function

Private Function ErrStr(rc As Long) As String
    Select Case rc
        Case X64_OK:          ErrStr = "OK"
        Case X64_ERR_NOINIT:  ErrStr = "ERR_NOINIT"
        Case X64_ERR_NODLL:   ErrStr = "ERR_NODLL"
        Case X64_ERR_NOPROC:  ErrStr = "ERR_NOPROC"
        Case X64_ERR_TOOMANY: ErrStr = "ERR_TOOMANY"
        Case X64_ERR_EXCEPT:  ErrStr = "ERR_EXCEPT"
        Case Else:            ErrStr = "UNKNOWN(" & rc & ")"
    End Select
End Function

' ============================================================
' STEP 0: Kiem tra phien ban DLL va trang thai khoi tao
' Chay dau tien de xac nhan DLL load duoc
' ============================================================
Public Sub Step0_CheckVersion()
    Dim ver As Long
    Dim hProc As Long, pCode As Long
    Dim state As Long
    Dim msg As String
    ' Kiem tra load DLL (X64Version khong can X64Init)
    ver = X64Version()
    Debug.Print ver
    msg = "X64Bridge.dll version: " & Hex$(ver) & vbCrLf
    ' Khoi tao
    If X64Init() = 0 Then
        msg = msg & "X64Init: THAT BAI" & vbCrLf
        Debug.Print msg, vbCritical, "Step0"
        Exit Sub
    End If
    msg = msg & "X64Init: OK" & vbCrLf
    ' Trang thai
    state = X64DebugInitState(hProc, pCode)
    msg = msg & "g_hProc OK : " & ((state And 1) <> 0) & " = " & Hex$(hProc) & vbCrLf
    msg = msg & "g_pCode OK : " & ((state And 2) <> 0) & " = " & Hex$(pCode) & vbCrLf
    msg = msg & "IsInit OK  : " & ((state And 4) <> 0) & vbCrLf
    X64Free
    Debug.Print msg, vbInformation, "Step0: Version & Init"
End Sub

' ============================================================
' STEP 1: Kiem tra NtWow64ReadVirtualMemory64
' Day la buoc nen tang — neu fail thi WOW64 trick khong chay duoc
' ============================================================
Public Sub Step1_CheckReadMem()
    Dim ntstatus As Long
    Dim rc As Long
    Dim msg As String
    X64Init
    rc = X64DebugReadMem(ntstatus)
    X64Free

    Select Case rc
        Case 1
            msg = "NtWow64QueryInformationProcess64 THAT BAI" & vbCrLf & "NTSTATUS = " & Hex$(ntstatus) & vbCrLf & vbCrLf & "Neu NTSTATUS = 80000002: ham nay khong ton tai" & vbCrLf & "-> Process khong phai WOW64 (dang chay trong host 64-bit?)"
            Debug.Print msg, vbCritical, "Step1: FAIL"
        Case 2
            Debug.Print "PebBaseAddress = 0 (bat thuong)", vbCritical, "Step1: FAIL"
        Case 3
            msg = "NtWow64ReadVirtualMemory64 THAT BAI" & vbCrLf & "NTSTATUS = " & Hex$(ntstatus)
            Debug.Print msg, vbCritical, "Step1: FAIL"
        Case Else
            msg = "NtWow64ReadVirtualMemory64: OK" & vbCrLf & "PEB64 low addr = " & Hex$(rc) & vbCrLf & "NTSTATUS = " & Hex$(ntstatus)
            Debug.Print msg, vbInformation, "Step1: OK"
    End Select
End Sub

' ============================================================
' STEP 2: Kiem tra GetModuleHandle64 va GetProcAddress64
' ============================================================
Public Sub Step2_CheckNtdll()
    Dim modLo As Long, modHi As Long
    Dim pfnLo As Long, pfnHi As Long
    Dim rc As Long
    Dim msg As String
    X64Init
    rc = X64DebugNtdll(modLo, modHi, pfnLo, pfnHi)
    X64Free

    Select Case rc
        Case 0
            msg = "ntdll.dll 64-bit base : " & Hex64(modLo, modHi) & vbCrLf & "NtGetCurProcessorNum  : " & Hex64(pfnLo, pfnHi) & vbCrLf & vbCrLf & "-> CA HAI OK"
            Debug.Print msg, vbInformation, "Step2: OK"
        Case 1
            Debug.Print "X64GetModuleHandle64 that bai (ntdll base = 0)", vbCritical, "Step2: FAIL"
        Case 2
            msg = "ntdll base: " & Hex64(modLo, modHi) & vbCrLf & "X64GetProcAddress64 that bai (ham = 0)"
            Debug.Print msg, vbCritical, "Step2: FAIL"
    End Select
End Sub

' ============================================================
' STEP 3: Kiem tra cache DLL (FindOrLoadW)
' ============================================================
Sub Step3_CheckCache()
    Dim modLo As Long
    Dim modHi As Long
    Dim cacheRc As Long
    Dim msg As String
    Call X64Init
    cacheRc = X64DebugCache(StrPtr("ntdll.dll"), modLo, modHi)
    Call X64Free

    If cacheRc = &HFFFF0001 Or cacheRc = -65535 Then
        Debug.Print "X64IsInitialized = False (init chua xong)"
        Exit Sub
    End If
    msg = ""
    msg = msg & "Return               : " & Hex$(cacheRc) & vbCrLf
    msg = msg & "Handle               : " & Hex64(modLo, modHi) & vbCrLf

    If (cacheRc And &H10000) <> 0 Then
        msg = msg & "X64GetModuleHandle64 : OK" & vbCrLf
    Else
        msg = msg & "X64GetModuleHandle64 : FAIL" & vbCrLf
    End If

    If (cacheRc And &H20000) <> 0 Then
        msg = msg & "FindOrLoadW          : OK" & vbCrLf
    Else
        msg = msg & "FindOrLoadW          : FAIL" & vbCrLf
    End If
    msg = msg & "g_CacheN sau goi     : " & CStr(cacheRc And &HFFFF&)

    If (cacheRc And &H30000) = &H30000 Then
        Debug.Print msg & vbCrLf & vbCrLf & "-> CA HAI OK"
    Else
        Debug.Print msg & vbCrLf & vbCrLf & "-> CO BUOC THAT BAI"
    End If
End Sub

' =========================================================
' TEST CACHE
' =========================================================
Public Sub Step3_TestCache()
    Dim hLo As Long
    Dim hHi As Long
    Dim rc As Long
    Dim msg As String
    Call X64Init
    rc = X64DebugCache(StrPtr("ntdll.dll"), hLo, hHi)
    Call X64Free
    msg = ""
    msg = msg & "Return : " & Hex$(rc) & vbCrLf
    msg = msg & "Handle : " & Hex64(hLo, hHi) & vbCrLf

    If (rc And &H10000) <> 0 Then
        msg = msg & "X64GetModuleHandle64 : OK" & vbCrLf
    Else
        msg = msg & "X64GetModuleHandle64 : FAIL" & vbCrLf
    End If

    If (rc And &H20000) <> 0 Then
        msg = msg & "FindOrLoadW : OK" & vbCrLf
    Else
        msg = msg & "FindOrLoadW : FAIL" & vbCrLf
    End If
    msg = msg & "g_CacheN : " & CStr(rc And &HFFFF&)
    Debug.Print msg
End Sub

' ============================================================
' STEP 4: Goi ham 64-bit thuc su — NtGetCurrentProcessorNumber
' Ham don gian nhat: 0 tham so, tra ve ULONG
' ============================================================
Public Sub Step4_CallNoArgs()
    Dim p(0) As X64Param   ' mang placeholder, nArgs=0 nen khong doc
    Dim rLo As Long, rHi As Long
    Dim rc As Long
    If X64Init() = 0 Then Debug.Print "X64Init that bai", vbCritical: Exit Sub
    rc = X64Invoke("ntdll.dll", "NtGetCurrentProcessorNumber", 0, p(0), rLo, rHi)
    X64Free

    If rc = X64_OK Then
        Debug.Print "NtGetCurrentProcessorNumber = " & rLo & vbCrLf & "(so thu tu CPU logic dang chay thread nay)", vbInformation, "Step4: OK"
    Else
        Debug.Print "X64Invoke that bai: " & ErrStr(rc), vbCritical, "Step4: FAIL"
    End If
End Sub

' ============================================================
' STEP 5: Goi ham 64-bit co 1 tham so output (pointer)
' NtQuerySystemTime(OUT PLARGE_INTEGER SystemTime)
' Tra ve Windows FILETIME 64-bit (100ns tu 1601-01-01)
' ============================================================
Public Sub Step5_CallWithOutputPtr()
    Dim p(0) As X64Param
    Dim rLo As Long, rHi As Long
    Dim rc As Long
    Dim bLo As Long, bHi As Long
    Dim ftLo As Long, ftHi As Long
    If X64Init() = 0 Then Debug.Print "X64Init that bai", vbCritical: Exit Sub
    ' Cap phat 8 byte chua LARGE_INTEGER output
    bLo = X64BufAllocLo(8)
    bHi = X64BufAllocHi(8)

    If bLo = 0 And bHi = 0 Then
        Debug.Print "BufAlloc that bai", vbCritical
X64Free:                                         Exit Sub
    End If
    ' Truyen dia chi buffer lam tham so
    p(0).lo = bLo: p(0).hi = bHi
    rc = X64Invoke("ntdll.dll", "NtQuerySystemTime", 1, p(0), rLo, rHi)

    If rc = X64_OK Then
        ftLo = X64BufReadDW(bLo, bHi, 0)   ' LARGE_INTEGER.LowPart
        ftHi = X64BufReadDW(bLo, bHi, 4)   ' LARGE_INTEGER.HighPart
        Debug.Print "NtQuerySystemTime OK" & vbCrLf & "NTSTATUS  = " & Hex$(rLo) & vbCrLf & "FILETIME  = " & Hex64(ftLo, ftHi) & vbCrLf & "(100ns units tu 1601-01-01)", vbInformation, "Step5: OK"
    Else
        Debug.Print "X64Invoke that bai: " & ErrStr(rc), vbCritical, "Step5: FAIL"
    End If
    X64BufFree bLo, bHi
    X64Free
End Sub

' ============================================================
' STEP 6: RtlGetVersion — 1 tham so input/output struct
' Lay phien ban Windows chinh xac (khong bi shim)
' RTL_OSVERSIONINFOW = 148 byte:
' +0  ULONG dwOSVersionInfoSize  = 148
' +4  ULONG dwMajorVersion
' +8  ULONG dwMinorVersion
' +12 ULONG dwBuildNumber
' ============================================================
Public Sub Step6_RtlGetVersion()
    Dim p(0) As X64Param
    Dim rLo As Long, rHi As Long
    Dim rc As Long
    Dim bLo As Long, bHi As Long
    Dim sHdr As String * 4
    If X64Init() = 0 Then Debug.Print "X64Init that bai", vbCritical: Exit Sub
    bLo = X64BufAllocLo(148)
    bHi = X64BufAllocHi(148)

    If bLo = 0 And bHi = 0 Then
        Debug.Print "BufAlloc that bai", vbCritical
X64Free:                                         Exit Sub
    End If
    ' Ghi dwOSVersionInfoSize = 148 tai offset 0 (little-endian 4 byte)
    Mid$(sHdr, 1, 1) = Chr$(148 And 255)
    Mid$(sHdr, 2, 1) = Chr$(0)
    Mid$(sHdr, 3, 1) = Chr$(0)
    Mid$(sHdr, 4, 1) = Chr$(0)
    X64BufWriteAnsi bLo, bHi, sHdr
    p(0).lo = bLo: p(0).hi = bHi
    rc = X64Invoke("ntdll.dll", "RtlGetVersion", 1, p(0), rLo, rHi)

    If rc = X64_OK And rLo = 0 Then   ' NTSTATUS = 0 = STATUS_SUCCESS
        Debug.Print "RtlGetVersion OK" & vbCrLf & "Major : " & X64BufReadDW(bLo, bHi, 4) & vbCrLf & "Minor : " & X64BufReadDW(bLo, bHi, 8) & vbCrLf & "Build : " & X64BufReadDW(bLo, bHi, 12) & vbCrLf & "(Win10=19xxx, Win11=22xxx)", vbInformation, "Step6: OK"
    Else
        Debug.Print "That bai: rc=" & ErrStr(rc) & " NTSTATUS=" & Hex$(rLo), vbCritical, "Step6: FAIL"
    End If
    X64BufFree bLo, bHi
    X64Free
End Sub

' ============================================================
' STEP 7: GetTickCount64 — 0 tham so, ket qua 64-bit
' Tra ve uptime Windows tinh bang millisecond
' ============================================================
Public Sub Step7_GetTickCount64()
    Dim p(0) As X64Param
    Dim rLo As Long, rHi As Long
    Dim rc As Long
    Dim ms As Currency
    Dim totalSec As Long
    If X64Init() = 0 Then Debug.Print "X64Init that bai", vbCritical: Exit Sub
    rc = X64Invoke("kernel32.dll", "GetTickCount64", 0, p(0), rLo, rHi)
    X64Free

    If rc <> X64_OK Then
        Debug.Print "X64Invoke that bai: " & ErrStr(rc), vbCritical, "Step7: FAIL"
        Exit Sub
    End If
    ' Ghep Lo+Hi thanh Currency de tinh (tranh overflow)
    ms = CCur(rHi And &H7FFFFFFF) * CCur(4294967296#)
    If (rHi And &H80000000) <> 0 Then ms = ms + CCur(2147483648#) * CCur(4294967296#)
    ms = ms + CCur(rLo And &H7FFFFFFF)
    If (rLo And &H80000000) <> 0 Then ms = ms + CCur(2147483648#)
    totalSec = CLng(ms / 1000)
    Debug.Print "GetTickCount64 = " & Hex64(rLo, rHi) & " ms" & vbCrLf & "Uptime: " & totalSec \ 3600 & " gio " & (totalSec Mod 3600) \ 60 & " phut " & totalSec Mod 60 & " giay", vbInformation, "Step7: OK"
End Sub

' ============================================================
' STEP 8: NtQueryPerformanceCounter — 2 tham so output
' Lay high-resolution timer (tuong duong QueryPerformanceCounter)
' ============================================================
Public Sub Step8_PerfCounter()
    Dim p(1) As X64Param
    Dim rLo As Long, rHi As Long
    Dim rc As Long
    Dim bcLo As Long, bcHi As Long   ' counter buffer
    Dim bfLo As Long, bfHi As Long   ' frequency buffer
    If X64Init() = 0 Then Debug.Print "X64Init that bai", vbCritical: Exit Sub
    bcLo = X64BufAllocLo(8): bcHi = X64BufAllocHi(8)
    bfLo = X64BufAllocLo(8): bfHi = X64BufAllocHi(8)
    p(0).lo = bcLo: p(0).hi = bcHi   ' &PerformanceCounter
    p(1).lo = bfLo: p(1).hi = bfHi   ' &PerformanceFrequency
    rc = X64Invoke("ntdll.dll", "NtQueryPerformanceCounter", 2, p(0), rLo, rHi)

    If rc = X64_OK Then
        Dim cLo As Long, cHi As Long
        Dim fLo As Long, fHi As Long
        cLo = X64BufReadDW(bcLo, bcHi, 0): cHi = X64BufReadDW(bcLo, bcHi, 4)
        fLo = X64BufReadDW(bfLo, bfHi, 0): fHi = X64BufReadDW(bfLo, bfHi, 4)
        Debug.Print "Counter   = " & Hex64(cLo, cHi) & vbCrLf & "Frequency = " & Hex64(fLo, fHi) & " tick/s", vbInformation, "Step8: OK"
    Else
        Debug.Print "X64Invoke that bai: " & ErrStr(rc), vbCritical, "Step8: FAIL"
    End If
    X64BufFree bcLo, bcHi
    X64BufFree bfLo, bfHi
    X64Free
End Sub

' ============================================================
' STEP 9: NtQueryInformationProcess — 5 tham so, lay PEB64 addr
' ============================================================
Public Sub Step9_GetPEB64()
    Dim p(4) As X64Param
    Dim rLo As Long, rHi As Long
    Dim rc As Long
    Dim bLo As Long, bHi As Long
    If X64Init() = 0 Then Debug.Print "X64Init that bai", vbCritical: Exit Sub
    ' PROCESS_BASIC_INFORMATION 64-bit = 48 byte
    ' PEB* tai offset +8
    bLo = X64BufAllocLo(48): bHi = X64BufAllocHi(48)
    ' NtQueryInformationProcess(-1, 0, &pbi, 48, NULL)
    ' hProcess = -1 = GetCurrentProcess pseudo-handle
    p(0).lo = &HFFFFFFFF: p(0).hi = &HFFFFFFFF  ' (HANDLE)-1
    p(1).lo = 0:          p(1).hi = 0            ' ProcessBasicInformation
    p(2).lo = bLo:        p(2).hi = bHi          ' &pbi
    p(3).lo = 48:         p(3).hi = 0            ' sizeof(pbi)
    p(4).lo = 0:          p(4).hi = 0            ' ReturnLength = NULL
    rc = X64Invoke("ntdll.dll", "NtQueryInformationProcess", 5, p(0), rLo, rHi)

    If rc = X64_OK And rLo = 0 Then   ' NTSTATUS = 0 = success
        ' PEB* tai offset +8 (64-bit pointer = 2 DWORD)
        Dim pebLo As Long, pebHi As Long
        pebLo = X64BufReadDW(bLo, bHi, 8)
        pebHi = X64BufReadDW(bLo, bHi, 12)
        Debug.Print "NtQueryInformationProcess OK" & vbCrLf & "PEB64 address = " & Hex64(pebLo, pebHi), vbInformation, "Step9: OK"
    Else
        Debug.Print "That bai: rc=" & ErrStr(rc) & " NTSTATUS=" & Hex$(rLo), vbCritical, "Step9: FAIL"
    End If
    X64BufFree bLo, bHi
    X64Free
End Sub

' ============================================================
' Chay tat ca theo thu tu — dung de test toan bo pipeline
' Dung lai neu co buoc FAIL de debug
' ============================================================
Public Sub RunAllSteps()
    Step0_CheckVersion
    Step1_CheckReadMem
    Step2_CheckNtdll
    Step3_CheckCache
    Step3_TestCache
    Step4_CallNoArgs
    Step5_CallWithOutputPtr
    Step6_RtlGetVersion
    Step7_GetTickCount64
    Step8_PerfCounter
    Step9_GetPEB64
    Debug.Print "Hoan thanh tat ca 11 buoc!", vbInformation, "X64Bridge VB6 Test"
End Sub
