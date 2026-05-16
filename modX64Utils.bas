Attribute VB_Name = "modX64Utils"
Option Explicit
' ==== Khai báo ki?u d? li?u ====
Private Type X64Param
    lo As Long
    hi As Long
End Type
Private Declare Function X64Init Lib "X64Bridge.dll" () As Long
Private Declare Sub X64Free Lib "X64Bridge.dll" ()
Private Declare Function X64Invoke Lib "X64Bridge.dll" (ByVal sDll As String, ByVal sFuncName As String, ByVal nArgs As Long, ByRef pParams As X64Param, ByRef rLo As Long, ByRef rHi As Long) As Long
Private Declare Function X64BufAllocLo Lib "X64Bridge.dll" (ByVal nBytes As Long) As Long
Private Declare Function X64BufAllocHi Lib "X64Bridge.dll" (ByVal nBytes As Long) As Long
Private Declare Sub X64BufFree Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long)
Private Declare Sub X64BufWriteAnsi Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long, ByVal s As String)
Private Declare Function X64BufReadDW Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long, ByVal ofs As Long) As Long
Private Const X64_OK As Long = 0
' ==== Hàm ti?n ích ====
' 1. L?y phiên b?n Windows
Public Function GetWindowsVersion() As String
    Dim p(0) As X64Param, rLo As Long, rHi As Long, rc As Long
    Dim bLo As Long, bHi As Long, major As Long, minor As Long, build As Long
    Dim sHdr As String * 4
    If X64Init() = 0 Then GetWindowsVersion = "Init Fail": Exit Function
    bLo = X64BufAllocLo(148): bHi = X64BufAllocHi(148)
    Mid$(sHdr, 1, 1) = Chr$(148 And 255)
    X64BufWriteAnsi bLo, bHi, sHdr
    p(0).lo = bLo: p(0).hi = bHi
    rc = X64Invoke("ntdll.dll", "RtlGetVersion", 1, p(0), rLo, rHi)

    If rc = X64_OK And rLo = 0 Then
        major = X64BufReadDW(bLo, bHi, 4)
        minor = X64BufReadDW(bLo, bHi, 8)
        build = X64BufReadDW(bLo, bHi, 12)
        GetWindowsVersion = "Windows " & major & "." & minor & " Build " & build
    Else
        GetWindowsVersion = "Error rc=" & rc & " NTSTATUS=" & Hex$(rLo)
    End If
    X64BufFree bLo, bHi
    X64Free
End Function

' 2. L?y uptime h? th?ng (giây)
Public Function GetSystemUptimeSeconds() As Long
    Dim p(0) As X64Param, rLo As Long, rHi As Long, rc As Long
    Dim ms As Currency
    If X64Init() = 0 Then GetSystemUptimeSeconds = -1: Exit Function
    rc = X64Invoke("kernel32.dll", "GetTickCount64", 0, p(0), rLo, rHi)
    X64Free
    If rc <> X64_OK Then GetSystemUptimeSeconds = -1: Exit Function
    ms = CCur(rHi And &H7FFFFFFF) * CCur(4294967296#)
    If (rHi And &H80000000) <> 0 Then ms = ms + CCur(2147483648#) * CCur(4294967296#)
    ms = ms + CCur(rLo And &H7FFFFFFF)
    If (rLo And &H80000000) <> 0 Then ms = ms + CCur(2147483648#)
    GetSystemUptimeSeconds = CLng(ms / 1000)
End Function

' 3. L?y th?i gian h? th?ng (FILETIME)
Public Function GetSystemTimeFileTime() As String
    Dim p(0) As X64Param, rLo As Long, rHi As Long, rc As Long
    Dim bLo As Long, bHi As Long, ftLo As Long, ftHi As Long
    If X64Init() = 0 Then GetSystemTimeFileTime = "Init Fail": Exit Function
    bLo = X64BufAllocLo(8): bHi = X64BufAllocHi(8)
    p(0).lo = bLo: p(0).hi = bHi
    rc = X64Invoke("ntdll.dll", "NtQuerySystemTime", 1, p(0), rLo, rHi)

    If rc = X64_OK Then
        ftLo = X64BufReadDW(bLo, bHi, 0)
        ftHi = X64BufReadDW(bLo, bHi, 4)
        GetSystemTimeFileTime = "FILETIME = " & Hex$(ftHi) & Hex$(ftLo)
    Else
        GetSystemTimeFileTime = "Error rc=" & rc & " NTSTATUS=" & Hex$(rLo)
    End If
    X64BufFree bLo, bHi
    X64Free
End Function

' ==== Sub Test t?ng h?p ====
Public Sub RunAllWrappers()
    Debug.Print GetWindowsVersion()
    Debug.Print "Uptime (s): " & GetSystemUptimeSeconds()
    Debug.Print GetSystemTimeFileTime()
    Debug.Print "Hoàn thành test các hàm tien ích!", vbInformation, "X64Bridge VB6/VBA Utils"
End Sub
