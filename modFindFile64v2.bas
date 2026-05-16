Attribute VB_Name = "modFindFile64v2"
' modFindFile64v2.bas
' FindFiles64 toi uu: resolve ham 1 lan, goi truc tiep qua X64CallDirect
' Ghi chu khong dau vi VBE bi loi font Unicode
'
' Nguyen nhan FindFiles64 v1 cham:
'   - Moi file goi Call64("FindNextFileW",...) = 1 lan sinh machine code
'   - 3527 file x 2 goi = ~7000 lan sinh x64 machine code
'   - Overhead chinh la: resolve ten ham + cap phat/giai phong string
'
' Giai phap v2:
'   - Resolve pfn (dia chi ham 64-bit) 1 lan truoc vong lap
'   - Goi X64CallDirect(pfn, args) thay vi Call64("kernel32","FindNextFileW",...)
'   - Tranh cap phat string "W:..." moi vong lap

Option Explicit

' Khai bao X64Bridge.dll
Private Declare Function X64Init Lib "X64Bridge.dll" () As Long
Private Declare Sub X64Free Lib "X64Bridge.dll" ()

' Resolve module + ham 64-bit (1 lan)
Private Declare Function X64ModLo Lib "X64Bridge.dll" (ByVal sName As String) As Long
Private Declare Function X64ModHi Lib "X64Bridge.dll" (ByVal sName As String) As Long
Private Declare Function X64ProcLo Lib "X64Bridge.dll" (ByVal hLo As Long, ByVal hHi As Long, ByVal sFunc As String) As Long
Private Declare Function X64ProcHi Lib "X64Bridge.dll" (ByVal hLo As Long, ByVal hHi As Long, ByVal sFunc As String) As Long

' Goi ham 64-bit voi dia chi da resolve (Lo/Hi rieng)
Private Declare Sub X64Call2 Lib "X64Bridge.dll" ( _
    ByVal fLo As Long, ByVal fHi As Long, _
    ByVal a1Lo As Long, ByVal a1Hi As Long, _
    ByVal a2Lo As Long, ByVal a2Hi As Long, _
    ByRef rLo As Long, ByRef rHi As Long)

Private Declare Sub X64Call1 Lib "X64Bridge.dll" ( _
    ByVal fLo As Long, ByVal fHi As Long, _
    ByVal a1Lo As Long, ByVal a1Hi As Long, _
    ByRef rLo As Long, ByRef rHi As Long)

' Buffer
Private Declare Function X64BufAllocLo Lib "X64Bridge.dll" (ByVal n As Long) As Long
Private Declare Function X64BufAllocHi Lib "X64Bridge.dll" (ByVal n As Long) As Long
Private Declare Sub X64BufFree Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long)
Private Declare Sub X64BufWriteUnicode Lib "X64Bridge.dll" ( _
    ByVal bLo As Long, ByVal bHi As Long, ByVal pW As Long)
Private Declare Function X64BufReadDW Lib "X64Bridge.dll" ( _
    ByVal bLo As Long, ByVal bHi As Long, ByVal ofs As Long) As Long

' Timer cao do
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Const WIN32_FIND_DATA_SIZE As Long = 592
Private Const OFS_ATTRIBUTES       As Long = 0
Private Const OFS_FILENAME         As Long = 44
Private Const FILE_ATTR_DIRECTORY  As Long = 16
Private Const INVALID_HANDLE_VALUE As Long = -1

' ============================================================
' Timer
' ============================================================
Private Function TimerMS() As Double
    Dim cnt As Currency, frq As Currency
    QueryPerformanceCounter cnt
    QueryPerformanceFrequency frq
    If frq = 0 Then Exit Function
    TimerMS = (cnt / frq) * 1000#
End Function

' ============================================================
' Doc ten file tu WIN32_FIND_DATAW (doc 4 byte = 2 WChar 1 lan)
' ============================================================
Private Function ReadFileName(bLo As Long, bHi As Long) As String
    Dim i   As Long
    Dim dw  As Long
    Dim c1  As Integer
    Dim c2  As Integer
    Dim s   As String
    s = ""
    For i = 0 To 129
        dw = X64BufReadDW(bLo, bHi, OFS_FILENAME + i * 4)
        c1 = dw And &HFFFF&
        c2 = (dw \ &H10000) And &HFFFF&
        If c1 = 0 Then Exit For
        s = s & Chr$(c1)
        If c2 = 0 Then Exit For
        s = s & Chr$(c2)
    Next i
    ReadFileName = s
End Function

' ============================================================
' FindFiles64v2 — toi uu: resolve pfn 1 lan, goi X64Call2 truc tiep
'
' So voi v1:
'   v1: moi file goi Call64("kernel32","FindNextFileW",...) → parse ten, alloc, free
'   v2: resolve pfnFirst/pfnNext/pfnClose 1 lan, goi X64Call2(pfn,...) trong vong lap
'       → bo toan bo overhead string parsing cua Call64
' ============================================================
Public Function FindFiles64v2(ByVal FolderPath As String, _
                               ByVal Ext As String, _
                               ByRef results() As String) As Long
    Dim sPattern  As String
    Dim hModLo    As Long, hModHi    As Long
    Dim pfnFstLo  As Long, pfnFstHi  As Long  ' FindFirstFileW
    Dim pfnNxtLo  As Long, pfnNxtHi  As Long  ' FindNextFileW
    Dim pfnClsLo  As Long, pfnClsHi  As Long  ' FindClose
    Dim pPatLo    As Long, pPatHi    As Long   ' buffer chua pattern string
    Dim pDatLo    As Long, pDatHi    As Long   ' WIN32_FIND_DATAW buffer
    Dim hFindLo   As Long, hFindHi   As Long   ' HANDLE
    Dim rLo       As Long, rHi       As Long
    Dim sName     As String
    Dim attrs     As Long
    Dim nFound    As Long
    Dim patBytes  As Long

    FindFiles64v2 = 0
    nFound = 0
    ReDim results(0)

    ' Xay dung pattern
    If Right$(FolderPath, 1) = "\" Then
        sPattern = FolderPath & Ext
    Else
        sPattern = FolderPath & "\" & Ext
    End If

    ' --- Buoc 1: resolve kernel32.dll 64-bit (1 lan) ---
    hModLo = X64ModLo("kernel32.dll")
    hModHi = X64ModHi("kernel32.dll")
    If hModLo = 0 And hModHi = 0 Then Exit Function

    ' --- Buoc 2: resolve 3 ham (1 lan) ---
    pfnFstLo = X64ProcLo(hModLo, hModHi, "FindFirstFileW")
    pfnFstHi = X64ProcHi(hModLo, hModHi, "FindFirstFileW")
    pfnNxtLo = X64ProcLo(hModLo, hModHi, "FindNextFileW")
    pfnNxtHi = X64ProcHi(hModLo, hModHi, "FindNextFileW")
    pfnClsLo = X64ProcLo(hModLo, hModHi, "FindClose")
    pfnClsHi = X64ProcHi(hModLo, hModHi, "FindClose")
    If pfnFstLo = 0 And pfnFstHi = 0 Then Exit Function

    ' --- Buoc 3: cap phat buffer pattern Unicode (1 lan) ---
    patBytes = (Len(sPattern) + 1) * 2
    pPatLo = X64BufAllocLo(patBytes)
    pPatHi = X64BufAllocHi(patBytes)
    X64BufWriteUnicode pPatLo, pPatHi, StrPtr(sPattern)

    ' --- Buoc 4: cap phat WIN32_FIND_DATAW buffer (1 lan) ---
    pDatLo = X64BufAllocLo(WIN32_FIND_DATA_SIZE)
    pDatHi = X64BufAllocHi(WIN32_FIND_DATA_SIZE)

    ' --- Buoc 5: FindFirstFileW(pPattern, pData) ---
    X64Call2 pfnFstLo, pfnFstHi, _
             pPatLo, pPatHi, _
             pDatLo, pDatHi, _
             hFindLo, hFindHi

    If hFindLo = INVALID_HANDLE_VALUE And hFindHi = -1 Then GoTo Cleanup
    If hFindLo = INVALID_HANDLE_VALUE And hFindHi = 0 Then GoTo Cleanup
    If hFindLo = 0 Then GoTo Cleanup

    ' --- Buoc 6: vong lap FindNextFileW (khong qua Call64 overhead) ---
    Do
        sName = ReadFileName(pDatLo, pDatHi)

        If sName <> "." And sName <> ".." Then
            attrs = X64BufReadDW(pDatLo, pDatHi, OFS_ATTRIBUTES)
            If (attrs And FILE_ATTR_DIRECTORY) = 0 Then
                If nFound > 0 Then ReDim Preserve results(nFound)
                results(nFound) = sName
                nFound = nFound + 1
            End If
        End If

        ' FindNextFileW(hFind, pData) — goi truc tiep, khong qua Call64
        X64Call2 pfnNxtLo, pfnNxtHi, _
                 hFindLo, 0, _
                 pDatLo, pDatHi, _
                 rLo, rHi

        If rLo = 0 Then Exit Do  ' FALSE = het file

    Loop

    ' FindClose(hFind)
    X64Call1 pfnClsLo, pfnClsHi, hFindLo, 0, rLo, rHi

Cleanup:
    If pPatLo <> 0 Then X64BufFree pPatLo, pPatHi
    If pDatLo <> 0 Then X64BufFree pDatLo, pDatHi

    If nFound = 0 Then ReDim results(0)
    FindFiles64v2 = nFound
End Function

' ============================================================
' FindFiles_VB6 — tim file bang Dir() thuan VB6 (tham chieu)
' ============================================================
Public Function FindFiles_VB6(ByVal FolderPath As String, _
                               ByVal Ext As String, _
                               ByRef results() As String) As Long
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
' Benchmark so sanh v1 / v2 / VB6
' ============================================================
Public Sub Benchmark_Compare(Optional ByVal FolderPath As String = "C:\Windows\System32", _
                               Optional ByVal Ext As String = "*.dll", _
                               Optional ByVal nRuns As Long = 5)
    Dim files() As String
    Dim i       As Long
    Dim t0      As Double, t1 As Double
    Dim n64v1   As Long, n64v2 As Long, nVB6 As Long
    Dim tot1    As Double, tot2 As Double, totVB6 As Double
    Dim bst1    As Double, bst2 As Double, bstVB6 As Double

    bst1 = 1E+15: bst2 = 1E+15: bstVB6 = 1E+15

    ' Warm up
    FindFiles64 FolderPath, Ext, files
    FindFiles64v2 FolderPath, Ext, files
    FindFiles_VB6 FolderPath, Ext, files

    X64Init

    ' --- v1 ---
    For i = 1 To nRuns
        t0 = TimerMS()
        n64v1 = FindFiles64(FolderPath, Ext, files)
        t1 = TimerMS()
        Dim r1 As Double: r1 = t1 - t0
        tot1 = tot1 + r1
        If r1 < bst1 Then bst1 = r1
        DoEvents
    Next i

    ' --- v2 ---
    For i = 1 To nRuns
        t0 = TimerMS()
        n64v2 = FindFiles64v2(FolderPath, Ext, files)
        t1 = TimerMS()
        Dim r2 As Double: r2 = t1 - t0
        tot2 = tot2 + r2
        If r2 < bst2 Then bst2 = r2
        DoEvents
    Next i

    ' --- VB6 ---
    For i = 1 To nRuns
        t0 = TimerMS()
        nVB6 = FindFiles_VB6(FolderPath, Ext, files)
        t1 = TimerMS()
        Dim rV As Double: rV = t1 - t0
        totVB6 = totVB6 + rV
        If rV < bstVB6 Then bstVB6 = rV
        DoEvents
    Next i

    X64Free

    Dim avg1  As Double: avg1 = tot1 / nRuns
    Dim avg2  As Double: avg2 = tot2 / nRuns
    Dim avgVB6 As Double: avgVB6 = totVB6 / nRuns

    Dim msg As String
    msg = "Benchmark: " & FolderPath & "\" & Ext & _
          "  (" & nRuns & " runs)" & vbCrLf & vbCrLf & _
          String$(56, "-") & vbCrLf & _
          "                  v1 (Call64)  v2 (Direct)  Dir() VB6" & vbCrLf & _
          String$(56, "-") & vbCrLf & _
          "So file tim duoc: " & _
              Format$(n64v1, "@@@@@@@") & "    " & _
              Format$(n64v2, "@@@@@@@") & "    " & _
              Format$(nVB6, "@@@@@@@") & vbCrLf & _
          "Trung binh (ms) : " & _
              Format$(avg1, "0000.000") & "    " & _
              Format$(avg2, "0000.000") & "    " & _
              Format$(avgVB6, "0000.000") & vbCrLf & _
          "Nhanh nhat (ms) : " & _
              Format$(bst1, "0000.000") & "    " & _
              Format$(bst2, "0000.000") & "    " & _
              Format$(bstVB6, "0000.000") & vbCrLf & _
          String$(56, "-") & vbCrLf & vbCrLf

    ' Toc do tuong doi so voi VB6
    If avg2 > 0 Then
        Dim ratio2 As Double: ratio2 = avgVB6 / avg2
        If ratio2 >= 1 Then
            msg = msg & "v2 cham hon Dir() " & Format$(ratio2, "0.0") & "x" & vbCrLf
        Else
            msg = msg & "v2 NHANH HON Dir() " & Format$(1 / ratio2, "0.0") & "x" & vbCrLf
        End If
    End If

    If avg1 > 0 Then
        msg = msg & "v2 nhanh hon v1: " & Format$(avg1 / avg2, "0.0") & "x" & vbCrLf
    End If

    msg = msg & vbCrLf & _
          "Ghi chu:" & vbCrLf & _
          "  v1 = Call64() trong vong lap (~7000 lan sinh machine code)" & vbCrLf & _
          "  v2 = X64Call2() truc tiep, resolve pfn 1 lan truoc vong lap" & vbCrLf & _
          "  VB6 = Dir() -> SysWOW64 (process 32-bit bi redirect)" & vbCrLf & _
          "  v2 thay System32 THAT (64-bit), Dir() thay SysWOW64"

    Debug.Print msg
    MsgBox msg, vbInformation, "Benchmark v1 vs v2 vs VB6"
End Sub


Sub Test_Benchmark_Compare()
Benchmark_Compare "C:\Windows\System32", "*.dll", 5
End Sub
