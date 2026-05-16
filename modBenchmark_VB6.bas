Attribute VB_Name = "modBenchmark_VB6"
Option Explicit
' =========================================================
' X64Bridge.dll benchmark for VB6
' =========================================================
Private Type X64Param
    lo As Long
    hi As Long
End Type
' =========================================================
' X64Bridge API
' =========================================================
Private Declare Function X64Init Lib "X64Bridge.dll" () As Long
Private Declare Sub X64Free Lib "X64Bridge.dll" ()
Private Declare Function X64Invoke Lib "X64Bridge.dll" (ByVal sDll As String, ByVal sFuncName As String, ByVal nArgs As Long, ByRef pParams As X64Param, ByRef rLo As Long, ByRef rHi As Long) As Long
Private Declare Function X64BufAllocLo Lib "X64Bridge.dll" (ByVal nBytes As Long) As Long
Private Declare Function X64BufAllocHi Lib "X64Bridge.dll" (ByVal nBytes As Long) As Long
Private Declare Sub X64BufFree Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long)
Private Declare Sub X64BufWriteUnicode Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long, ByVal pW As Long)
Private Declare Function X64BufReadDW Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long, ByVal ofs As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Const INVALID_HANDLE_VALUE As Long = -1
' =========================================================
' Read Unicode string from bridge buffer
' =========================================================
Private Function ReadUnicode(ByVal bLo As Long, ByVal bHi As Long, Optional ByVal maxChars As Long = 260) As String
    Dim i As Long
    Dim ch As Integer
    Dim s As String

    For i = 0 To maxChars - 1
        ch = X64BufReadDW(bLo, bHi, i * 2) And &HFFFF&
        If ch = 0 Then Exit For
        s = s & ChrW$(ch)
    Next
    ReadUnicode = s
End Function

' =========================================================
' VB6 benchmark
' =========================================================
Public Sub Benchmark_VB6()
    Dim t1 As Long
    Dim t2 As Long
    Dim f As String
    Dim n As Long
    Debug.Print String(60, "=")
    Debug.Print "VB6 Dir$ benchmark"
    Debug.Print String(60, "=")
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

' =========================================================
' 64-bit benchmark
' =========================================================
Public Sub Benchmark_X64()
    Dim rc As Long
    Dim t1 As Long
    Dim t2 As Long
    Dim args(1) As X64Param
    Dim retLo As Long
    Dim retHi As Long
    Dim pathLo As Long
    Dim pathHi As Long
    Dim dataLo As Long
    Dim dataHi As Long
    Dim hFindLo As Long
    Dim hFindHi As Long
    Dim fileName As String
    Debug.Print String(60, "=")
    Debug.Print "X64 FindFirstFileW benchmark"
    Debug.Print String(60, "=")

    If X64Init() = 0 Then
        Debug.Print "X64Init FAIL"
        Exit Sub
    End If
    ' WIN32_FIND_DATAW ~ 592 bytes
    dataLo = X64BufAllocLo(592)
    dataHi = X64BufAllocHi(592)
    pathLo = X64BufAllocLo(1024)
    pathHi = X64BufAllocHi(1024)
    Call X64BufWriteUnicode(pathLo, pathHi, StrPtr("C:\Windows\*.*"))
    args(0).lo = pathLo
    args(0).hi = pathHi
    args(1).lo = dataLo
    args(1).hi = dataHi
    t1 = GetTickCount()
    rc = X64Invoke("kernel32.dll", "FindFirstFileW", 2, args(0), retLo, retHi)
    t2 = GetTickCount()

    If rc <> 0 Then
        Debug.Print "X64Invoke FAIL : "; rc
    Else
        hFindLo = retLo
        hFindHi = retHi
        Debug.Print "Find handle : 0x" & Right$("00000000" & Hex$(hFindHi), 8) & Right$("00000000" & Hex$(hFindLo), 8)
        Debug.Print "Time (ms)   : "; (t2 - t1)

        If hFindLo <> INVALID_HANDLE_VALUE Then
            Debug.Print "SUCCESS CALL 64-bit API"
            ' cFileName offset = 44
            fileName = ReadUnicode(dataLo, dataHi, 260)
            Debug.Print "First file  : "; fileName
        Else
            Debug.Print "FindFirstFileW FAIL"
        End If
    End If
    Call X64BufFree(pathLo, pathHi)
    Call X64BufFree(dataLo, dataHi)
    Call X64Free
End Sub

' =========================================================
' Run all
' =========================================================
Public Sub RunAll2()
    Debug.Print
    Debug.Print "========== X64Bridge Benchmark =========="
    Debug.Print
    Call Benchmark_VB6
    Debug.Print
    Call Benchmark_X64
    Debug.Print
    Debug.Print "Done."
End Sub
