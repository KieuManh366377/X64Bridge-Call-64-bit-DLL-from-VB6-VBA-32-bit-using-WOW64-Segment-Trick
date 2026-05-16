Attribute VB_Name = "modBenchmark64"
' modBenchmark64.bas
' So sanh toc do: FindFiles64 (64-bit API) vs Dir() thuan VB6
' Ghi chu khong dau vi VBE bi loi font Unicode
'
' Yeu cau: modCall64.bas va modFindFile64.bas trong cung project
Option Explicit
Private Declare Function X64BufReadDW Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long, ByVal ofs As Long) As Long
Private Declare Function X64BufAllocLo Lib "X64Bridge.dll" (ByVal n As Long) As Long
Private Declare Function X64BufAllocHi Lib "X64Bridge.dll" (ByVal n As Long) As Long
Private Declare Sub X64BufFree Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long)
' ============================================================
' Timer cao do dung QueryPerformanceCounter (32-bit API)
' ============================================================
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Function TimerMS() As Double
    Dim cnt As Currency
    Dim frq As Currency
    QueryPerformanceCounter cnt
    QueryPerformanceFrequency frq
    If frq = 0 Then TimerMS = 0: Exit Function
    TimerMS = (cnt / frq) * 1000#   ' millisecond
End Function

' ============================================================
' [A] Tim file bang Dir() thuan VB6 32-bit
' ============================================================
Public Function FindFiles_VB6(ByVal FolderPath As String, ByVal Ext As String, ByRef results() As String) As Long
    Dim sPattern As String
    Dim sName    As String
    Dim nFound   As Long

    If Right$(FolderPath, 1) = "\" Then
        sPattern = FolderPath & Ext
    Else
        sPattern = FolderPath & "\" & Ext
    End If
    nFound = 0
    ReDim results(0)
    sName = Dir$(sPattern, vbNormal)

    Do While Len(sName) > 0
        If nFound > 0 Then ReDim Preserve results(nFound)
        results(nFound) = sName
        nFound = nFound + 1
        sName = Dir$()
    Loop
    If nFound = 0 Then ReDim results(0)
    FindFiles_VB6 = nFound
End Function

' ============================================================
' [B] Benchmark chinh — chay N lan moi phuong phap
' ============================================================
Public Sub Benchmark_FindFiles(Optional ByVal FolderPath As String = "C:\Windows\System32", Optional ByVal Ext As String = "*.dll", Optional ByVal nRuns As Long = 5)
    Dim i       As Long
    Dim n64     As Long
    Dim nVB6    As Long
    Dim Files() As String
    Dim t0      As Double
    Dim t1      As Double
    Dim total64 As Double
    Dim totalVB6 As Double
    Dim best64  As Double
    Dim bestVB6 As Double
    Dim run64   As Double
    Dim runVB6  As Double
    Dim msg     As String
    best64 = 1E+15
    bestVB6 = 1E+15
    total64 = 0
    totalVB6 = 0
    ' --- Warm up: chay 1 lan truoc de OS cache file list ---
    FindFiles64 FolderPath, Ext, Files
    FindFiles_VB6 FolderPath, Ext, Files
    ' ── Benchmark FindFiles64 (64-bit API) ──────────────────
    For i = 1 To nRuns
        t0 = TimerMS()
        n64 = FindFiles64(FolderPath, Ext, Files)
        t1 = TimerMS()
        run64 = t1 - t0
        total64 = total64 + run64
        If run64 < best64 Then best64 = run64
        DoEvents
    Next i
    ' ── Benchmark FindFiles_VB6 (Dir() thuan) ───────────────
    For i = 1 To nRuns
        t0 = TimerMS()
        nVB6 = FindFiles_VB6(FolderPath, Ext, Files)
        t1 = TimerMS()
        runVB6 = t1 - t0
        totalVB6 = totalVB6 + runVB6
        If runVB6 < bestVB6 Then bestVB6 = runVB6
        DoEvents
    Next i
    ' ── Ket qua ─────────────────────────────────────────────
    Dim avg64   As Double
    Dim avgVB6  As Double
    avg64 = total64 / nRuns
    avgVB6 = totalVB6 / nRuns
    Dim ratio As Double
    If avg64 > 0 Then ratio = avgVB6 / avg64
    msg = "Benchmark: FindFirstFileW 64-bit  vs  Dir() VB6" & vbCrLf & "Thu muc : " & FolderPath & "\" & Ext & vbCrLf & "So lan  : " & nRuns & vbCrLf & vbCrLf
    msg = msg & String$(48, "-") & vbCrLf
    msg = msg & "                  FindFiles64    Dir() VB6" & vbCrLf
    msg = msg & String$(48, "-") & vbCrLf
    msg = msg & "So file tim duoc: " & Format$(n64, "@@@@@@@@@@") & "    " & Format$(nVB6, "@@@@@@@@@@") & vbCrLf
    msg = msg & "Trung binh (ms) : " & Format$(avg64, "0.000") & "         " & Format$(avgVB6, "0.000") & vbCrLf
    msg = msg & "Nhanh nhat (ms) : " & Format$(best64, "0.000") & "         " & Format$(bestVB6, "0.000") & vbCrLf
    msg = msg & String$(48, "-") & vbCrLf & vbCrLf

    If ratio >= 1 Then
        msg = msg & ">> Dir() VB6 nhanh hon " & Format$(ratio, "0.00") & "x so voi FindFiles64" & vbCrLf
        msg = msg & "   (ly do: Dir() goi Win32 API 32-bit, khong qua WOW64 trick)"
    ElseIf ratio > 0 Then
        msg = msg & ">> FindFiles64 nhanh hon " & Format$(1 / ratio, "0.00") & "x so voi Dir() VB6" & vbCrLf
        msg = msg & "   (bat ngo: overhead WOW64 thap hon du kien)"
    End If
    Debug.Print msg
    ' MsgBox msg, vbInformation, "Benchmark FindFiles"
    Call64_Free
End Sub

' ============================================================
' [C] Benchmark chi tiet: do tung buoc trong FindFiles64
' de biet buoc nao chiem nhieu thoi gian nhat
' ============================================================
Public Sub Benchmark_Detail(Optional ByVal FolderPath As String = "C:\Windows\System32", Optional ByVal Ext As String = "*.dll")
    Dim sPattern As String
    Dim t0       As Double
    Dim t1       As Double
    Dim tAlloc   As Double
    Dim tFirst   As Double
    Dim tLoop    As Double
    Dim tFree    As Double
    Dim nFiles   As Long
    Dim r        As String
    Dim hFind    As Long
    Dim bLo      As Long
    Dim bHi      As Long
    Dim sName    As String
    Dim attrs    As Long
    Dim msg      As String
    Const WIN32_FIND_DATA_SIZE As Long = 592
    Const OFS_ATTRIBUTES       As Long = 0
    Const OFS_FILENAME         As Long = 44
    Const FILE_ATTR_DIRECTORY  As Long = 16
    Const INVALID_HANDLE       As Long = -1

    If Right$(FolderPath, 1) = "\" Then
        sPattern = FolderPath & Ext
    Else
        sPattern = FolderPath & "\" & Ext
    End If
    ' --- Do: cap phat buffer ---
    t0 = TimerMS()
    bLo = X64BufAllocLo(WIN32_FIND_DATA_SIZE)
    bHi = X64BufAllocHi(WIN32_FIND_DATA_SIZE)
    t1 = TimerMS()
    tAlloc = t1 - t0
    ' --- Do: FindFirstFileW ---
    t0 = TimerMS()
    r = Call64("kernel32.dll", "FindFirstFileW", "W:" & sPattern, CLng(bLo))
    t1 = TimerMS()
    tFirst = t1 - t0

    If Not IsRetOK(r) Then
        MsgBox "FindFirstFileW that bai: " & r, vbCritical
        X64BufFree bLo, bHi
        Exit Sub
    End If
    hFind = ParseRetLo(r)

    If hFind = INVALID_HANDLE Or hFind = 0 Then
        MsgBox "Khong tim thay file nao", vbInformation
        X64BufFree bLo, bHi
        Exit Sub
    End If
    ' --- Do: toan bo vong lap FindNextFileW ---
    nFiles = 0
    t0 = TimerMS()

    Do
        ' Doc ten (bo qua . va ..)
        Dim i As Long, Lo As Long, c1 As Integer, c2 As Integer
        sName = ""

        For i = 0 To 129
            Lo = X64BufReadDW(bLo, bHi, OFS_FILENAME + i * 4)
            c1 = Lo And &HFFFF&
            c2 = (Lo \ &H10000) And &HFFFF&
            If c1 = 0 Then Exit For
            sName = sName & Chr$(c1)
            If c2 = 0 Then Exit For
            sName = sName & Chr$(c2)
        Next i

        If sName <> "." And sName <> ".." Then
            attrs = X64BufReadDW(bLo, bHi, OFS_ATTRIBUTES)

            If (attrs And FILE_ATTR_DIRECTORY) = 0 Then
                nFiles = nFiles + 1
            End If
        End If
        r = Call64("kernel32.dll", "FindNextFileW", hFind, CLng(bLo))
        If Not IsRetOK(r) Then Exit Do
        If ParseRetLo(r) = 0 Then Exit Do
    Loop
    t1 = TimerMS()
    tLoop = t1 - t0
    ' --- Do: giai phong ---
    t0 = TimerMS()
    Call64 "kernel32.dll", "FindClose", hFind
    X64BufFree bLo, bHi
    t1 = TimerMS()
    tFree = t1 - t0
    ' --- Ket qua ---
    Dim tTotal As Double
    tTotal = tAlloc + tFirst + tLoop + tFree
    msg = "Benchmark chi tiet FindFiles64:" & vbCrLf & FolderPath & "\" & Ext & vbCrLf & "So file tim duoc: " & nFiles & vbCrLf & vbCrLf & String$(40, "-") & vbCrLf & "BufAlloc        : " & _
        Format$(tAlloc, "0.000") & " ms" & Format$(tAlloc / tTotal * 100, "  0.0") & "%" & vbCrLf & "FindFirstFileW  : " & Format$(tFirst, "0.000") & " ms" & Format$(tFirst / tTotal * 100, "  0.0") & _
        "%" & vbCrLf & "FindNextFileW x" & nFiles & ": " & Format$(tLoop, "0.000") & " ms" & Format$(tLoop / tTotal * 100, "  0.0") & "%" & vbCrLf & "FindClose+Free  : " & Format$(tFree, "0.000") & _
        " ms" & Format$(tFree / tTotal * 100, "  0.0") & "%" & vbCrLf & String$(40, "-") & vbCrLf & "Tong cong       : " & Format$(tTotal, "0.000") & " ms" & vbCrLf & vbCrLf & "Trung binh moi file: " & _
        Format$(tLoop / IIf(nFiles > 0, nFiles, 1), "0.000") & " ms/file"
    Debug.Print msg
    ' MsgBox msg, vbInformation, "Benchmark Detail"
    Call64_Free
End Sub

Sub Test_Benchmark_FindFiles()
    ' So sanh tong the (5 lan do moi phuong phap)
    ' Benchmark_FindFiles "C:\Windows\System32", "*.dll", 5
    ' Chi tiet tung buoc trong FindFiles64
    ' Benchmark_Detail "C:\Windows\System32", "*.dll"
    ' Goi truc tiep
    Dim Files() As String
    Dim i, n As Long
    n = FindFiles64("C:\Windows\System32", "*.dll", Files)

    For i = 0 To n - 1
        Debug.Print Files(i)
    Next i
End Sub
