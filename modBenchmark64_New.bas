Attribute VB_Name = "modBenchmark64_New"
' =========================================================
' modBenchmark64_New.bas
' Benchmark FindFirstFileW 64-bit qua X64Bridge.dll
'
' REQUIRE:
'   - X64Bridge.dll
'   - modCall64.bas
'
' TEST:
'   Call Test_Benchmark64
'
' =========================================================

Option Explicit

' =========================================================
' X64Bridge API
' =========================================================

Private Type X64Param
    lo As Long
    hi As Long
End Type

Private Declare Function X64Init Lib "X64Bridge.dll" () As Long
Private Declare Sub X64Free Lib "X64Bridge.dll" ()

Private Declare Function X64Invoke Lib "X64Bridge.dll" ( _
    ByVal sDll As String, _
    ByVal sFunc As String, _
    ByVal nArgs As Long, _
    ByRef pArgs As X64Param, _
    ByRef rLo As Long, _
    ByRef rHi As Long) As Long

Private Declare Function X64BufAllocLo Lib "X64Bridge.dll" ( _
    ByVal nBytes As Long) As Long

Private Declare Function X64BufAllocHi Lib "X64Bridge.dll" ( _
    ByVal nBytes As Long) As Long

Private Declare Sub X64BufFree Lib "X64Bridge.dll" ( _
    ByVal bLo As Long, _
    ByVal bHi As Long)

Private Declare Sub X64BufWriteUnicode Lib "X64Bridge.dll" ( _
    ByVal bLo As Long, _
    ByVal bHi As Long, _
    ByVal pW As Long)

Private Declare Function X64BufReadDW Lib "X64Bridge.dll" ( _
    ByVal bLo As Long, _
    ByVal bHi As Long, _
    ByVal ofs As Long) As Long

' =========================================================
' TIMER
' =========================================================

Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
    lpPerformanceCount As Currency) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
    lpFrequency As Currency) As Long

Private Function TimerMS() As Double

    Dim c As Currency
    Dim f As Currency

    QueryPerformanceCounter c
    QueryPerformanceFrequency f

    If f = 0 Then
        TimerMS = 0
    Else
        TimerMS = (c / f) * 1000#
    End If

End Function

' =========================================================
' READ UNICODE FROM X64 BUFFER
' =========================================================

Private Function ReadUnicode64( _
    ByVal bLo As Long, _
    ByVal bHi As Long, _
    ByVal ofs As Long, _
    Optional ByVal maxChars As Long = 260) As String

    Dim i As Long
    Dim dw As Long
    Dim c1 As Integer
    Dim c2 As Integer
    Dim s As String

    s = ""

    For i = 0 To maxChars - 1 Step 2

        dw = X64BufReadDW(bLo, bHi, ofs + (i * 2))

        c1 = dw And &HFFFF&
        c2 = (dw \ &H10000) And &HFFFF&

        If c1 = 0 Then Exit For
        s = s & ChrW$(c1)

        If c2 = 0 Then Exit For
        s = s & ChrW$(c2)

    Next i

    ReadUnicode64 = s

End Function

' =========================================================
' CALL 64-BIT API
' =========================================================

Private Function Call64_2( _
    ByVal sDll As String, _
    ByVal sFunc As String, _
    ByVal p1Lo As Long, _
    ByVal p1Hi As Long, _
    ByVal p2Lo As Long, _
    ByVal p2Hi As Long, _
    ByRef rLo As Long, _
    ByRef rHi As Long) As Long

    Dim a(1) As X64Param

    a(0).lo = p1Lo
    a(0).hi = p1Hi

    a(1).lo = p2Lo
    a(1).hi = p2Hi

    Call64_2 = X64Invoke( _
                    sDll, _
                    sFunc, _
                    2, _
                    a(0), _
                    rLo, _
                    rHi)

End Function

Private Function Call64_1( _
    ByVal sDll As String, _
    ByVal sFunc As String, _
    ByVal p1Lo As Long, _
    ByVal p1Hi As Long, _
    ByRef rLo As Long, _
    ByRef rHi As Long) As Long

    Dim a(0) As X64Param

    a(0).lo = p1Lo
    a(0).hi = p1Hi

    Call64_1 = X64Invoke( _
                    sDll, _
                    sFunc, _
                    1, _
                    a(0), _
                    rLo, _
                    rHi)

End Function

' =========================================================
' FIND FILES 64-BIT
' =========================================================

Private Function FindFiles64( _
    ByVal FolderPath As String, _
    ByVal Mask As String, _
    ByRef Files() As String) As Long

    Const WIN32_FIND_DATA_SIZE As Long = 592
    Const OFS_FILENAME As Long = 44

    Dim path As String

    Dim pathLo As Long
    Dim pathHi As Long

    Dim dataLo As Long
    Dim dataHi As Long

    Dim retLo As Long
    Dim retHi As Long

    Dim rc As Long

    Dim hFind As Long

    Dim sName As String
    Dim n As Long

    If Right$(FolderPath, 1) = "\" Then
        path = FolderPath & Mask
    Else
        path = FolderPath & "\" & Mask
    End If

    ReDim Files(0)

    rc = X64Init()

    If rc = 0 Then
        Debug.Print "X64Init FAIL"
        Exit Function
    End If

    pathLo = X64BufAllocLo(1024)
    pathHi = X64BufAllocHi(1024)

    dataLo = X64BufAllocLo(WIN32_FIND_DATA_SIZE)
    dataHi = X64BufAllocHi(WIN32_FIND_DATA_SIZE)

    X64BufWriteUnicode pathLo, pathHi, StrPtr(path)

    rc = Call64_2( _
            "kernel32.dll", _
            "FindFirstFileW", _
            pathLo, pathHi, _
            dataLo, dataHi, _
            retLo, retHi)

    If rc <> 0 Then
        Debug.Print "FindFirstFileW invoke FAIL : "; rc
        GoTo CLEANUP
    End If

    hFind = retLo

    If hFind = -1 Or hFind = 0 Then
        Debug.Print "No files found"
        GoTo CLEANUP
    End If

    Do

        sName = ReadUnicode64(dataLo, dataHi, OFS_FILENAME)

        If sName <> "." And sName <> ".." Then

            If n > 0 Then
                ReDim Preserve Files(n)
            End If

            Files(n) = sName
            n = n + 1

        End If

        rc = Call64_2( _
                "kernel32.dll", _
                "FindNextFileW", _
                hFind, 0, _
                dataLo, dataHi, _
                retLo, retHi)

        If rc <> 0 Then Exit Do
        If retLo = 0 Then Exit Do

    Loop

    Call64_1 _
        "kernel32.dll", _
        "FindClose", _
        hFind, 0, _
        retLo, retHi

    FindFiles64 = n

CLEANUP:

    X64BufFree pathLo, pathHi
    X64BufFree dataLo, dataHi

    X64Free

End Function

' =========================================================
' VB6 DIR BENCHMARK
' =========================================================

Public Function FindFilesVB6( _
    ByVal FolderPath As String, _
    ByVal Mask As String, _
    ByRef Files() As String) As Long

    Dim s As String
    Dim path As String
    Dim n As Long

    If Right$(FolderPath, 1) = "\" Then
        path = FolderPath & Mask
    Else
        path = FolderPath & "\" & Mask
    End If

    ReDim Files(0)

    s = Dir$(path)

    Do While LenB(s) <> 0

        If n > 0 Then
            ReDim Preserve Files(n)
        End If

        Files(n) = s

        n = n + 1

        s = Dir$

    Loop

    FindFilesVB6 = n

End Function

' =========================================================
' MAIN BENCHMARK
' =========================================================

Public Sub Test_Benchmark64()

    Dim arr() As String

    Dim t1 As Double
    Dim t2 As Double

    Dim n As Long

    Debug.Print String$(60, "=")
    Debug.Print "VB6 Dir$ Benchmark"
    Debug.Print String$(60, "=")

    t1 = TimerMS()

    n = FindFilesVB6( _
            "C:\Windows\System32", _
            "*.dll", _
            arr)

    t2 = TimerMS()

    Debug.Print "Files : "; n
    Debug.Print "Time  : "; Format$(t2 - t1, "0.000"); " ms"

    Debug.Print

    Debug.Print String$(60, "=")
    Debug.Print "X64 FindFirstFileW Benchmark"
    Debug.Print String$(60, "=")

    t1 = TimerMS()

    n = FindFiles64( _
            "C:\Windows\System32", _
            "*.dll", _
            arr)

    t2 = TimerMS()

    Debug.Print "Files : "; n
    Debug.Print "Time  : "; Format$(t2 - t1, "0.000"); " ms"

    Debug.Print

    If n > 0 Then
        Debug.Print "First file : "; arr(0)
    End If

    Debug.Print
    Debug.Print "DONE"

End Sub

