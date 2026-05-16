Attribute VB_Name = "modCall64"
' modCall64.bas
' Wrapper trung gian: Call64() goi bat ky ham 64-bit nao qua X64Bridge.dll
' VB6 32-bit, Windows 64-bit (WOW64)
' Ghi chu khong dau vi VBE bi loi font Unicode
'
' CACH DUNG:
' Dim r As String
' r = Call64("ntdll.dll", "NtGetCurrentProcessorNumber")
' r = Call64("ntdll.dll", "RtlGetVersion", "O:148")
' r = Call64("kernel32.dll", "GetTickCount64")
'
' DINH DANG THAM SO (ParamArray):
' Long / Integer  -> truyen truc tiep (Lo = gia tri, Hi = 0)
' "O:N"           -> Output buffer N byte, tu dong cap phat,
' ket qua doc lai sau khi goi
' "W:text"        -> Unicode string (cap phat buffer, ghi WChar)
' "A:text"        -> ANSI string (cap phat buffer, ghi AnsiChar)
' Null / Empty    -> NULL pointer (Lo=0, Hi=0)
'
' GIA TRI TRA VE:
' "OK:HHH:LLL"    -> thanh cong, HHH=rHi, LLL=rLo (hex 8 chu so)
' "ERR:ma_loi"    -> that bai
' Dung ParseRet() de lay gia tri
Option Explicit
' ============================================================
' Kieu du lieu noi bo
' ============================================================
Private Type X64Param
    Lo As Long
    Hi As Long
End Type
' ============================================================
' Khai bao X64Bridge.dll
' ============================================================
Private Declare Function X64Init Lib "X64Bridge.dll" () As Long
Private Declare Sub X64Free Lib "X64Bridge.dll" ()
Private Declare Function X64Invoke Lib "X64Bridge.dll" (ByVal sDll As String, ByVal sFuncName As String, ByVal nArgs As Long, ByRef pParams As X64Param, ByRef rLo As Long, ByRef rHi As Long) As Long
Private Declare Function X64BufAllocLo Lib "X64Bridge.dll" (ByVal nBytes As Long) As Long
Private Declare Function X64BufAllocHi Lib "X64Bridge.dll" (ByVal nBytes As Long) As Long
Private Declare Sub X64BufFree Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long)
Private Declare Sub X64BufWriteAnsi Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long, ByVal s As String)
Private Declare Sub X64BufWriteUnicode Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long, ByVal pW As Long)
Private Declare Function X64BufReadDW Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long, ByVal ofs As Long) As Long
' Ma loi
Private Const X64_OK          As Long = 0
Private Const X64_ERR_NOINIT  As Long = 1
Private Const X64_ERR_NODLL   As Long = 2
Private Const X64_ERR_NOPROC  As Long = 3
Private Const X64_ERR_TOOMANY As Long = 4
Private Const X64_ERR_EXCEPT  As Long = 5
' Loai tham so noi bo
Private Const PT_VAL   As Long = 0  ' so nguyen truyen truc tiep
Private Const PT_OUT   As Long = 1  ' output buffer "O:N"
Private Const PT_WSTR  As Long = 2  ' unicode string "W:text"
Private Const PT_ASTR  As Long = 3  ' ansi string "A:text"
Private Const PT_NULL  As Long = 4  ' null pointer
' So tham so toi da
Private Const MAX_ARGS As Long = 16
' ============================================================
' Bien trang thai module — khoi tao 1 lan, giu xuyen session
' ============================================================
Private g_bInit As Boolean
Private Sub EnsureInit()
    If Not g_bInit Then
        If X64Init() <> 0 Then g_bInit = True
    End If
End Sub

' Goi tuong minh de reset (dung khi can reinit)
Public Sub Call64_Free()
    If g_bInit Then
        X64Free
        g_bInit = False
    End If
End Sub

' ============================================================
' Helper noi bo
' ============================================================
Private Function Hex8(v As Long) As String
    Hex8 = Right$("00000000" & Hex$(v), 8)
End Function

Private Function Hex64(Lo As Long, Hi As Long) As String
    If Hi = 0 Then
        Hex64 = "0x" & Hex8(Lo)
    Else
        Hex64 = "0x" & Hex8(Hi) & Hex8(Lo)
    End If
End Function

Private Function MakeOK(rLo As Long, rHi As Long) As String
    MakeOK = "OK:" & Hex8(rHi) & ":" & Hex8(rLo)
End Function

Private Function MakeERR(sReason As String) As String
    MakeERR = "ERR:" & sReason
End Function

' ============================================================
' ParseRet — lay gia tri tu ket qua Call64
' ParseRet("OK:00000000:0000000A")  -> 10
' ParseRetHi("OK:00000001:00000000") -> 1
' ============================================================
Public Function ParseRetLo(sRet As String) As Long
    ' "OK:HHHHHHHH:LLLLLLLL" -> lay 8 chu so cuoi
    If Left$(sRet, 3) <> "OK:" Then ParseRetLo = 0: Exit Function
    ParseRetLo = CLng("&H" & Right$(sRet, 8))
End Function

Public Function ParseRetHi(sRet As String) As Long
    If Left$(sRet, 3) <> "OK:" Then ParseRetHi = 0: Exit Function
    ' lay 8 chu so giua: vi tri 4..11
    ParseRetHi = CLng("&H" & Mid$(sRet, 4, 8))
End Function

Public Function IsRetOK(sRet As String) As Boolean
    IsRetOK = (Left$(sRet, 3) = "OK:")
End Function

' ============================================================
' BufRead — doc DWORD tu output buffer sau khi Call64 tra ve
' Luu buffer handle trong bien ngoai va doc bang ham nay
' ============================================================
Public Function BufReadDW(bLo As Long, bHi As Long, offset As Long) As Long
    BufReadDW = X64BufReadDW(bLo, bHi, offset)
End Function

Public Sub BufFree(bLo As Long, bHi As Long)
    X64BufFree bLo, bHi
End Sub

' ============================================================
' Call64 — ham wrapper chinh
'
' sDll    : ten hoac duong dan DLL ("ntdll.dll", "C:\my.dll")
' sFunc   : ten ham export ("NtGetCurrentProcessorNumber")
' args()  : tham so dong (ParamArray), xem dinh dang o dau file
'
' Tra ve : "OK:HI:LO" hoac "ERR:ly_do"
'
' Vi du:
' Dim r As String
' r = Call64("ntdll.dll", "NtGetCurrentProcessorNumber")
' If IsRetOK(r) Then Debug.Print "CPU = " & ParseRetLo(r)
' ============================================================
Public Function Call64(ByVal sDll As String, ByVal sFunc As String, ParamArray args() As Variant) As String
    Dim p(MAX_ARGS - 1) As X64Param  ' mang tham so
    Dim nArgs As Long                 ' so tham so thuc su
    ' Luu buffer da cap phat de giai phong sau khi goi
    Dim bufLo(MAX_ARGS - 1) As Long
    Dim bufHi(MAX_ARGS - 1) As Long
    Dim bufSz(MAX_ARGS - 1) As Long   ' kich thuoc buffer (de biet la buffer hay khong)
    Dim bufTyp(MAX_ARGS - 1) As Long  ' loai tham so
    Dim rLo As Long, rHi As Long
    Dim rc  As Long
    Dim i   As Long
    Dim s   As String
    Dim nBufBytes As Long
    EnsureInit

    If Not g_bInit Then
        Call64 = MakeERR("NOINIT")
        Exit Function
    End If
    ' --- Xu ly tham so ---
    nArgs = 0

    For i = 0 To UBound(args)

        If nArgs >= MAX_ARGS Then
            Call64 = MakeERR("TOOMANY")
            GoTo Cleanup
        End If

        Select Case VarType(args(i))
        Case vbNull, vbEmpty
            ' NULL pointer
            bufTyp(nArgs) = PT_NULL
            p(nArgs).Lo = 0
            p(nArgs).Hi = 0
        Case vbString
            s = CStr(args(i))

            If Left$(s, 2) = "O:" Then
                ' --- Output buffer ---
                nBufBytes = CLng(Mid$(s, 3))
                If nBufBytes <= 0 Then nBufBytes = 8
                bufLo(nArgs) = X64BufAllocLo(nBufBytes)
                bufHi(nArgs) = X64BufAllocHi(nBufBytes)
                bufSz(nArgs) = nBufBytes
                bufTyp(nArgs) = PT_OUT
                p(nArgs).Lo = bufLo(nArgs)
                p(nArgs).Hi = bufHi(nArgs)
            ElseIf Left$(s, 2) = "W:" Then
                ' --- Unicode string ---
                Dim wStr As String
                wStr = Mid$(s, 3)
                Dim wBytes As Long
                wBytes = (Len(wStr) + 1) * 2
                bufLo(nArgs) = X64BufAllocLo(wBytes)
                bufHi(nArgs) = X64BufAllocHi(wBytes)
                bufSz(nArgs) = wBytes
                bufTyp(nArgs) = PT_WSTR
                X64BufWriteUnicode bufLo(nArgs), bufHi(nArgs), StrPtr(wStr)
                p(nArgs).Lo = bufLo(nArgs)
                p(nArgs).Hi = bufHi(nArgs)
            ElseIf Left$(s, 2) = "A:" Then
                ' --- ANSI string ---
                Dim aStr As String
                aStr = Mid$(s, 3)
                Dim aBytes As Long
                aBytes = Len(aStr) + 1
                bufLo(nArgs) = X64BufAllocLo(aBytes)
                bufHi(nArgs) = X64BufAllocHi(aBytes)
                bufSz(nArgs) = aBytes
                bufTyp(nArgs) = PT_ASTR
                X64BufWriteAnsi bufLo(nArgs), bufHi(nArgs), aStr
                p(nArgs).Lo = bufLo(nArgs)
                p(nArgs).Hi = bufHi(nArgs)
            Else
                ' ANSI string mac dinh (khong prefix)
                Dim dStr As String
                dStr = s
                Dim dBytes As Long
                dBytes = Len(dStr) + 1
                bufLo(nArgs) = X64BufAllocLo(dBytes)
                bufHi(nArgs) = X64BufAllocHi(dBytes)
                bufSz(nArgs) = dBytes
                bufTyp(nArgs) = PT_ASTR
                X64BufWriteAnsi bufLo(nArgs), bufHi(nArgs), dStr
                p(nArgs).Lo = bufLo(nArgs)
                p(nArgs).Hi = bufHi(nArgs)
            End If
        Case vbLong, vbInteger, vbByte
            ' So nguyen 32-bit: truyen truc tiep
            ' Sign-extend: neu am (bit31=1) thi Hi = &HFFFFFFFF
            ' Vi du: -1 = 0xFFFFFFFFFFFFFFFF, khong phai 0x00000000FFFFFFFF
            bufTyp(nArgs) = PT_VAL
            Dim lv1 As Long: lv1 = CLng(args(i))
            p(nArgs).Lo = lv1
            p(nArgs).Hi = IIf(lv1 < 0, &HFFFFFFFF, 0)
        Case vbBoolean
            bufTyp(nArgs) = PT_VAL
            p(nArgs).Lo = IIf(CBool(args(i)), 1, 0)
            p(nArgs).Hi = 0
        Case Else
            ' Co gang ep kieu Long, sign-extend tuong tu
            On Error Resume Next
            bufTyp(nArgs) = PT_VAL
            Dim lve As Long: lve = CLng(args(i))
            p(nArgs).Lo = lve
            p(nArgs).Hi = IIf(lve < 0, &HFFFFFFFF, 0)
            On Error GoTo 0
        End Select
        nArgs = nArgs + 1
    Next i
    ' --- Goi ham 64-bit ---
    rLo = 0: rHi = 0

    If nArgs = 0 Then
        ' Khong tham so: truyen dummy (nArgs=0, khong doc p(0))
        rc = X64Invoke(sDll, sFunc, 0, p(0), rLo, rHi)
    Else
        rc = X64Invoke(sDll, sFunc, nArgs, p(0), rLo, rHi)
    End If
    ' --- Xu ly ket qua ---
    Select Case rc
    Case X64_OK:          Call64 = MakeOK(rLo, rHi)
    Case X64_ERR_NOINIT:  Call64 = MakeERR("NOINIT")
    Case X64_ERR_NODLL:   Call64 = MakeERR("NODLL:" & sDll)
    Case X64_ERR_NOPROC:  Call64 = MakeERR("NOPROC:" & sFunc)
    Case X64_ERR_TOOMANY: Call64 = MakeERR("TOOMANY")
    Case X64_ERR_EXCEPT:  Call64 = MakeERR("EXCEPT")
    Case Else:            Call64 = MakeERR("RC=" & rc)
    End Select
Cleanup:
    ' --- Giai phong buffer da cap phat ---
    For i = 0 To nArgs - 1

        If bufSz(i) > 0 Then
            X64BufFree bufLo(i), bufHi(i)
        End If
    Next i
End Function

' ============================================================
' Call64Ex — giong Call64 nhung giu buffer output lai
' de nguoi goi tu doc ket qua struct sau khi ham chay xong
'
' Tra ve: "OK:HI:LO" hoac "ERR:..."
' pBufLo/pBufHi: nhan dia chi buffer output (loai "O:N")
' chi nhan buffer cua tham so output DAU TIEN (index 0 trong output list)
' nguoi goi phai goi BufFree(pBufLo, pBufHi) sau khi doc xong
'
' Vi du: RtlGetVersion
' Dim bLo As Long, bHi As Long
' Dim r As String
' r = Call64Ex("ntdll.dll", "RtlGetVersion", bLo, bHi, "O:148")
' If IsRetOK(r) Then
' Debug.Print "Build = " & BufReadDW(bLo, bHi, 12)
' BufFree bLo, bHi
' End If
' ============================================================
Public Function Call64Ex(ByVal sDll As String, ByVal sFunc As String, ByRef pBufLo As Long, ByRef pBufHi As Long, ParamArray args() As Variant) As String
    Dim p(MAX_ARGS - 1) As X64Param
    Dim nArgs As Long
    Dim bufLo(MAX_ARGS - 1) As Long
    Dim bufHi(MAX_ARGS - 1) As Long
    Dim bufSz(MAX_ARGS - 1) As Long
    Dim bufTyp(MAX_ARGS - 1) As Long
    Dim firstOutIdx As Long   ' index cua output buffer dau tien
    Dim rLo As Long, rHi As Long
    Dim rc  As Long
    Dim i   As Long
    Dim s   As String
    Dim nBufBytes As Long
    pBufLo = 0: pBufHi = 0
    firstOutIdx = -1
    EnsureInit

    If Not g_bInit Then
        Call64Ex = MakeERR("NOINIT")
        Exit Function
    End If
    nArgs = 0

    For i = 0 To UBound(args)

        If nArgs >= MAX_ARGS Then
            Call64Ex = MakeERR("TOOMANY")
            GoTo CleanupEx
        End If

        Select Case VarType(args(i))
        Case vbNull, vbEmpty
            bufTyp(nArgs) = PT_NULL
            p(nArgs).Lo = 0: p(nArgs).Hi = 0
        Case vbString
            s = CStr(args(i))

            If Left$(s, 2) = "O:" Then
                nBufBytes = CLng(Mid$(s, 3))
                If nBufBytes <= 0 Then nBufBytes = 8
                bufLo(nArgs) = X64BufAllocLo(nBufBytes)
                bufHi(nArgs) = X64BufAllocHi(nBufBytes)
                bufSz(nArgs) = nBufBytes
                bufTyp(nArgs) = PT_OUT
                p(nArgs).Lo = bufLo(nArgs)
                p(nArgs).Hi = bufHi(nArgs)
                ' Ghi size vao 4 byte dau (cho struct co cbSize)
                Dim sSzHdr As String * 4
                Mid$(sSzHdr, 1, 1) = Chr$(nBufBytes And 255)
                Mid$(sSzHdr, 2, 1) = Chr$((nBufBytes \ 256) And 255)
                Mid$(sSzHdr, 3, 1) = Chr$(0)
                Mid$(sSzHdr, 4, 1) = Chr$(0)
                X64BufWriteAnsi bufLo(nArgs), bufHi(nArgs), sSzHdr
                ' Luu buffer output dau tien de tra ve nguoi goi
                If firstOutIdx = -1 Then
                    firstOutIdx = nArgs
                    pBufLo = bufLo(nArgs)
                    pBufHi = bufHi(nArgs)
                End If
            ElseIf Left$(s, 2) = "W:" Then
                Dim wStr2 As String
                wStr2 = Mid$(s, 3)
                Dim wBytes2 As Long
                wBytes2 = (Len(wStr2) + 1) * 2
                bufLo(nArgs) = X64BufAllocLo(wBytes2)
                bufHi(nArgs) = X64BufAllocHi(wBytes2)
                bufSz(nArgs) = wBytes2
                bufTyp(nArgs) = PT_WSTR
                X64BufWriteUnicode bufLo(nArgs), bufHi(nArgs), StrPtr(wStr2)
                p(nArgs).Lo = bufLo(nArgs)
                p(nArgs).Hi = bufHi(nArgs)
            ElseIf Left$(s, 2) = "A:" Then
                Dim aStr2 As String
                aStr2 = Mid$(s, 3)
                bufLo(nArgs) = X64BufAllocLo(Len(aStr2) + 1)
                bufHi(nArgs) = X64BufAllocHi(Len(aStr2) + 1)
                bufSz(nArgs) = Len(aStr2) + 1
                bufTyp(nArgs) = PT_ASTR
                X64BufWriteAnsi bufLo(nArgs), bufHi(nArgs), aStr2
                p(nArgs).Lo = bufLo(nArgs)
                p(nArgs).Hi = bufHi(nArgs)
            Else
                Dim dStr2 As String
                dStr2 = s
                bufLo(nArgs) = X64BufAllocLo(Len(dStr2) + 1)
                bufHi(nArgs) = X64BufAllocHi(Len(dStr2) + 1)
                bufSz(nArgs) = Len(dStr2) + 1
                bufTyp(nArgs) = PT_ASTR
                X64BufWriteAnsi bufLo(nArgs), bufHi(nArgs), dStr2
                p(nArgs).Lo = bufLo(nArgs)
                p(nArgs).Hi = bufHi(nArgs)
            End If
        Case vbLong, vbInteger, vbByte, vbBoolean
            ' Sign-extend: -1 -> Hi=&HFFFFFFFF, gia tri duong -> Hi=0
            bufTyp(nArgs) = PT_VAL
            Dim lv2 As Long: lv2 = CLng(args(i))
            p(nArgs).Lo = lv2
            p(nArgs).Hi = IIf(lv2 < 0, &HFFFFFFFF, 0)
        Case Else
            On Error Resume Next
            bufTyp(nArgs) = PT_VAL
            Dim lve2 As Long: lve2 = CLng(args(i))
            p(nArgs).Lo = lve2
            p(nArgs).Hi = IIf(lve2 < 0, &HFFFFFFFF, 0)
            On Error GoTo 0
        End Select
        nArgs = nArgs + 1
    Next i
    rLo = 0: rHi = 0
    rc = X64Invoke(sDll, sFunc, nArgs, p(0), rLo, rHi)

    Select Case rc
    Case X64_OK:          Call64Ex = MakeOK(rLo, rHi)
    Case X64_ERR_NOINIT:  Call64Ex = MakeERR("NOINIT")
    Case X64_ERR_NODLL:   Call64Ex = MakeERR("NODLL:" & sDll)
    Case X64_ERR_NOPROC:  Call64Ex = MakeERR("NOPROC:" & sFunc)
    Case X64_ERR_TOOMANY: Call64Ex = MakeERR("TOOMANY")
    Case X64_ERR_EXCEPT:  Call64Ex = MakeERR("EXCEPT")
    Case Else:            Call64Ex = MakeERR("RC=" & rc)
    End Select
CleanupEx:
    ' Giai phong tat ca buffer TRU buffer output dau tien
    ' (nguoi goi chiu trach nhiem giai phong buffer do)
    For i = 0 To nArgs - 1

        If bufSz(i) > 0 And i <> firstOutIdx Then
            X64BufFree bufLo(i), bufHi(i)
        End If
    Next i
End Function

' ============================================================
' Demo / Test
' ============================================================
Public Sub Demo_Call64()
    Dim r As String
    Dim msg As String
    ' --- 1. Ham khong tham so ---
    r = Call64("ntdll.dll", "NtGetCurrentProcessorNumber")
    msg = "NtGetCurrentProcessorNumber:" & vbCrLf

    If IsRetOK(r) Then
        msg = msg & "  CPU index = " & ParseRetLo(r) & vbCrLf
    Else
        msg = msg & "  " & r & vbCrLf
    End If
    ' --- 2. GetTickCount64 (ket qua 64-bit) ---
    r = Call64("kernel32.dll", "GetTickCount64")
    msg = msg & "GetTickCount64:" & vbCrLf

    If IsRetOK(r) Then
        Dim lo64 As Long, hi64 As Long
        lo64 = ParseRetLo(r)
        hi64 = ParseRetHi(r)
        Dim ms As Currency
        ms = CCur(hi64 And &H7FFFFFFF) * CCur(4294967296#)
        If (hi64 And &H80000000) <> 0 Then ms = ms + CCur(2147483648#) * CCur(4294967296#)
        ms = ms + CCur(lo64 And &H7FFFFFFF)
        If (lo64 And &H80000000) <> 0 Then ms = ms + CCur(2147483648#)
        Dim sec As Long
        sec = CLng(ms / 1000)
        msg = msg & "  Uptime = " & sec \ 3600 & "h " & (sec Mod 3600) \ 60 & "m " & sec Mod 60 & "s" & vbCrLf
    Else
        msg = msg & "  " & r & vbCrLf
    End If
    ' --- 3. RtlGetVersion (dung Call64Ex giu buffer) ---
    Dim bLo As Long, bHi As Long
    r = Call64Ex("ntdll.dll", "RtlGetVersion", bLo, bHi, "O:148")
    msg = msg & "RtlGetVersion:" & vbCrLf

    If IsRetOK(r) And ParseRetLo(r) = 0 Then  ' NTSTATUS = 0
        msg = msg & "  Windows " & BufReadDW(bLo, bHi, 4) & "." & BufReadDW(bLo, bHi, 8) & " Build " & BufReadDW(bLo, bHi, 12) & vbCrLf
        BufFree bLo, bHi
    Else
        msg = msg & "  " & r & vbCrLf
    End If
    ' --- 4. NtQuerySystemTime (1 output pointer) ---
    Dim ftLo As Long, ftHi As Long
    r = Call64Ex("ntdll.dll", "NtQuerySystemTime", bLo, bHi, "O:8")
    msg = msg & "NtQuerySystemTime:" & vbCrLf

    If IsRetOK(r) Then
        ftLo = BufReadDW(bLo, bHi, 0)
        ftHi = BufReadDW(bLo, bHi, 4)
        msg = msg & "  FILETIME = 0x" & Right$("00000000" & Hex$(ftHi), 8) & Right$("00000000" & Hex$(ftLo), 8) & vbCrLf
        BufFree bLo, bHi
    Else
        msg = msg & "  " & r & vbCrLf
    End If
    ' --- 5. NtQueryInformationProcess 5 tham so ---
    ' hProcess=-1, class=0, &pbi, size=48, NULL
    r = Call64Ex("ntdll.dll", "NtQueryInformationProcess", bLo, bHi, CLng(-1), 0&, "O:48", 48&, 0&)
    msg = msg & "NtQueryInformationProcess (PEB64):" & vbCrLf

    If IsRetOK(r) And ParseRetLo(r) = 0 Then
        Dim pebLo As Long, pebHi As Long
        pebLo = BufReadDW(bLo, bHi, 8)
        pebHi = BufReadDW(bLo, bHi, 12)
        msg = msg & "  PEB64 = 0x" & Right$("00000000" & Hex$(pebHi), 8) & Right$("00000000" & Hex$(pebLo), 8) & vbCrLf
        BufFree bLo, bHi
    Else
        msg = msg & "  " & r & vbCrLf
    End If
    Debug.Print msg
    ' MsgBox msg, vbInformation, "Call64 Demo"
    Call64_Free
End Sub


' Test NtQueryInformationProcess dung Step9 style (da hoat dong)
Public Sub Demo_PEB64_Direct()
    Dim p(4) As X64Param
    Dim rLo As Long, rHi As Long
    Dim bLo As Long, bHi As Long
    X64Init
    bLo = X64BufAllocLo(48): bHi = X64BufAllocHi(48)
    ' hProcess = -1: Lo=&HFFFFFFFF, Hi=&HFFFFFFFF (sign-extended 64-bit)
    p(0).Lo = &HFFFFFFFF: p(0).Hi = &HFFFFFFFF
    p(1).Lo = 0:          p(1).Hi = 0
    p(2).Lo = bLo:        p(2).Hi = bHi
    p(3).Lo = 48:         p(3).Hi = 0
    p(4).Lo = 0:          p(4).Hi = 0
    Dim rc As Long
    rc = X64Invoke("ntdll.dll", "NtQueryInformationProcess", 5, p(0), rLo, rHi)

    If rc = 0 And rLo = 0 Then
        Dim pebLo As Long, pebHi As Long
        pebLo = X64BufReadDW(bLo, bHi, 8)
        pebHi = X64BufReadDW(bLo, bHi, 12)
        Debug.Print "PEB64 = 0x" & Right$("00000000" & Hex$(pebHi), 8) & Right$("00000000" & Hex$(pebLo), 8)
    Else
        Debug.Print "FAIL rc=" & rc & " NTSTATUS=" & Hex$(rLo)
    End If
    X64BufFree bLo, bHi
    X64Free
End Sub
