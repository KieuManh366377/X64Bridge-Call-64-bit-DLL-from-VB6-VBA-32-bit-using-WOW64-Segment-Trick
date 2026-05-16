Attribute VB_Name = "modFindFile64Test"
Option Explicit
' ============================================================
' X64Bridge Benchmark
'
' Test:
' 1. VB6 Dir$
' 2. 64-bit FindFirstFileW via X64Bridge.dll
'
' Target:
' VB6 Win32 + WOW64
'
' Author:
' Kieu Manh
' ============================================================
' ============================================================
' TYPES
' ============================================================
Private Type X64Param
    lo As Long
    hi As Long
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATAW
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName(0 To 259) As Integer
    cAlternate(0 To 13) As Integer
End Type
' ============================================================
' X64Bridge
' ============================================================
' ============================================================
' X64Bridge
' ============================================================
Private Declare Function X64Init Lib "X64Bridge.dll" () As Long
Private Declare Sub X64Free Lib "X64Bridge.dll" ()
Private Declare Function X64Invoke Lib "X64Bridge.dll" (ByVal sDll As String, ByVal sFuncName As String, ByVal nArgs As Long, ByRef pParams As X64Param, ByRef rLo As Long, ByRef rHi As Long) As Long
Private Declare Function X64BufAllocLo Lib "X64Bridge.dll" (ByVal nBytes As Long) As Long
Private Declare Function X64BufAllocHi Lib "X64Bridge.dll" (ByVal nBytes As Long) As Long
Private Declare Sub X64BufFree Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long)
Private Declare Sub X64BufWriteUnicode Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long, ByVal pW As Long)
Private Declare Function X64BufReadDW Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long, ByVal ofs As Long) As Long
' ============================================================
' 32-bit APIs
' ============================================================
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
' ============================================================
' CONSTANTS
' ============================================================
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const X64_OK As Long = 0
' ============================================================
' HELPERS
' ============================================================
Private Function Hex64(ByVal lo As Long, ByVal hi As Long) As String
    Hex64 = "0x" & Right$("00000000" & Hex$(hi), 8) & Right$("00000000" & Hex$(lo), 8)
End Function

Private Function UnicodeFromBuffer(ByVal bLo As Long, ByVal bHi As Long, ByVal offset As Long, ByVal maxChars As Long) As String
    Dim i As Long
    Dim ch As Integer
    Dim s As String
    s = ""

    For i = 0 To maxChars - 1
        ch = X64BufReadDW(bLo, bHi, offset + (i * 2)) And &HFFFF&
        If ch = 0 Then Exit For
        s = s & ChrW$(ch)
    Next
    UnicodeFromBuffer = s
End Function

' ============================================================
' TEST 1
' VB6 DIR$
' ============================================================
Public Sub Benchmark_VB6()
    Dim t1 As Long
    Dim t2 As Long
    Dim f As String
    Dim n As Long
    Debug.Print String$(60, "=")
    Debug.Print "VB6 Dir$ benchmark"
    Debug.Print String$(60, "=")
    t1 = GetTickCount()
    f = Dir$("C:\Windows\*.*")

    Do While LenB(f) <> 0
        n = n + 1
        f = Dir$
    Loop
    t2 = GetTickCount()
    Debug.Print "Files found : "; n
    Debug.Print "Time (ms)   : "; (t2 - t1)
End Sub

' ============================================================
' TEST 2
' 64-bit FindFirstFileW
' ============================================================
Public Sub Benchmark_X64()
    Dim rc As Long
    Dim t1 As Long
    Dim t2 As Long
    Dim args(1) As X64Param
    Dim retLo As Long
    Dim retHi As Long
    Dim hFindLo As Long
    Dim hFindHi As Long
    Dim pathLo As Long
    Dim pathHi As Long
    Dim dataLo As Long
    Dim dataHi As Long
    Dim firstName As String
    Debug.Print String$(60, "=")
    Debug.Print "X64 FindFirstFileW benchmark"
    Debug.Print String$(60, "=")
    rc = X64Init()

    If rc = 0 Then
        Debug.Print "X64Init FAIL"
        Exit Sub
    End If
    ' WIN32_FIND_DATAW ˜ 592 bytes
    dataLo = X64BufAllocLo(592)
    dataHi = X64BufAllocHi(592)
    pathLo = X64BufAllocLo(1024)
    pathHi = X64BufAllocHi(1024)

    If pathLo = 0 Then
        Debug.Print "Buffer alloc FAIL"
        GoTo CLEANUP
    End If
    ' write unicode path
    Call X64BufWriteUnicode(pathLo, pathHi, StrPtr("C:\Windows\*.*"))
    ' arg0 = LPCWSTR
    args(0).lo = pathLo
    args(0).hi = pathHi
    ' arg1 = LPWIN32_FIND_DATAW
    args(1).lo = dataLo
    args(1).hi = dataHi
    t1 = GetTickCount()
    rc = X64Invoke("kernel32.dll", "FindFirstFileW", 2, args(0), retLo, retHi)
    t2 = GetTickCount()

    If rc <> X64_OK Then
        Debug.Print "X64Invoke FAIL : "; rc
        GoTo CLEANUP
    End If
    hFindLo = retLo
    hFindHi = retHi
    Debug.Print "Find handle : "; Hex64(hFindLo, hFindHi)
    Debug.Print "Time (ms)   : "; (t2 - t1)

    If hFindLo = INVALID_HANDLE_VALUE Then
        Debug.Print "FindFirstFileW FAIL"
    Else
        Debug.Print "SUCCESS CALL 64-bit API"
        ' cFileName offset = 44
        firstName = UnicodeFromBuffer(dataLo, dataHi, 44, 260)
        Debug.Print "First file  : "; firstName
        Call FindClose(hFindLo)
    End If
CLEANUP:

    If pathLo <> 0 Then
        Call X64BufFree(pathLo, pathHi)
    End If

    If dataLo <> 0 Then
        Call X64BufFree(dataLo, dataHi)
    End If
    Call X64Free
End Sub

' ============================================================
' RUN ALL
' ============================================================
Public Sub RunAll()
    Debug.Print vbCrLf
    Debug.Print "========== X64Bridge Benchmark =========="
    Debug.Print
    Call Benchmark_VB6
    Debug.Print
    Call Benchmark_X64
    Debug.Print
    Debug.Print "Done."
End Sub
