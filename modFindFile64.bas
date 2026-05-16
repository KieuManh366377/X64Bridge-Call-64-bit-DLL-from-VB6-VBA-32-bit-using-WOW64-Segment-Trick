Attribute VB_Name = "modFindFile64"
' modFindFile64.bas
' Tim file trong thu muc bang FindFirstFileW / FindNextFileW 64-bit
' Goi qua Call64Ex (modCall64.bas phai duoc them vao project)
' Ghi chu khong dau vi VBE bi loi font Unicode

Option Explicit

' Khai bao ham X64BufReadDW de dung noi bo
Private Declare Function X64BufReadDW Lib "X64Bridge.dll" ( _
    ByVal bLo As Long, ByVal bHi As Long, ByVal ofs As Long) As Long

' Type noi bo tranh xung dot
Rem Private Type X64Param_
Private Type X64Param_
    lo As Long
    hi As Long
End Type

Private Declare Function X64BufAllocLo Lib "X64Bridge.dll" (ByVal n As Long) As Long
Private Declare Function X64BufAllocHi Lib "X64Bridge.dll" (ByVal n As Long) As Long
Private Declare Sub X64BufFree Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long)
' ============================================================
' WIN32_FIND_DATAW layout (592 byte tong)
' +0    DWORD  dwFileAttributes
' +4    FILETIME ftCreationTime       (8 byte)
' +12   FILETIME ftLastAccessTime     (8 byte)
' +20   FILETIME ftLastWriteTime      (8 byte)
' +28   DWORD  nFileSizeHigh
' +32   DWORD  nFileSizeLow
' +36   DWORD  dwReserved0
' +40   DWORD  dwReserved1
' +44   WCHAR  cFileName[260]         (520 byte)
' +564  WCHAR  cAlternateFileName[14] (28 byte)
' Tong = 592 byte
' ============================================================
Private Const WIN32_FIND_DATA_SIZE As Long = 592
Private Const OFS_ATTRIBUTES       As Long = 0
Private Const OFS_SIZE_HIGH        As Long = 28
Private Const OFS_SIZE_LOW         As Long = 32
Private Const OFS_FILENAME         As Long = 44   ' WCHAR[260]
Private Const INVALID_HANDLE       As Long = -1   ' INVALID_HANDLE_VALUE
Private Const FILE_ATTR_DIRECTORY  As Long = 16

' ============================================================
' Doc chuoi Unicode tu buffer (WCHAR array)
' Doc tung WORD (2 byte) den khi gap null terminator
' ============================================================
Private Function BufReadWStr(bLo As Long, bHi As Long, _
                              offset As Long, maxChars As Long) As String
    Dim i    As Long
    Dim w    As Long
    Dim sOut As String
    sOut = ""
    For i = 0 To maxChars - 1
        w = X64BufReadDW(bLo, bHi, offset + i * 2) And &HFFFF&
        ' Chi lay WORD thap (BufReadDW tra DWORD, lay 2 byte thap)
        ' Thuc ra moi lan doc 4 byte nhung chi dung 2 byte dau
        If w = 0 Then Exit For
        sOut = sOut & Chr$(w)
    Next i
    BufReadWStr = sOut
End Function


' ============================================================
' ReadFileName — doc ten file tu WIN32_FIND_DATAW trong buffer
' Doc 2 byte mot lan (moi WORD = 1 WChar)
' ============================================================
Private Function ReadFileName(bLo As Long, bHi As Long) As String
    Dim i    As Long
    Dim lo   As Long   ' DWORD doc duoc
    Dim hi   As Long
    Dim c1   As Integer, c2 As Integer
    Dim sOut As String
    sOut = ""

    ' cFileName tai +44, moi lan doc 4 byte = 2 WChar
    For i = 0 To 129   ' 260 WChar / 2 = 130 cap
        lo = X64BufReadDW(bLo, bHi, OFS_FILENAME + i * 4)
        c1 = lo And &HFFFF&           ' WChar thu 1 (byte thap)
        c2 = (lo \ &H10000) And &HFFFF& ' WChar thu 2 (byte cao)
        If c1 = 0 Then Exit For
        sOut = sOut & Chr$(c1)
        If c2 = 0 Then Exit For
        sOut = sOut & Chr$(c2)
    Next i
    ReadFileName = sOut
End Function

' ============================================================
' FindFiles64
' Tim tat ca file trong FolderPath khop voi Ext
'
' FolderPath : thu muc can tim, vi du "C:\Windows"
' Ext        : phan mo rong, vi du "*.exe", "*.dll", "*.*"
' results()  : mang String nhan ket qua (ten file)
' Tra ve     : so file tim duoc, -1 neu loi
'
' Vi du:
'   Dim files() As String
'   Dim n As Long
'   n = FindFiles64("C:\Windows", "*.exe", files)
'   For i = 0 To n-1 : Debug.Print files(i) : Next
' ============================================================
Public Function FindFiles64(ByVal FolderPath As String, _
                             ByVal Ext As String, _
                             ByRef results() As String) As Long
    Dim sPattern As String
    Dim hFind    As Long    ' HANDLE (32-bit handle trong WOW64 < 4GB)
    Dim bLo      As Long    ' WIN32_FIND_DATAW buffer
    Dim bHi      As Long
    Dim r        As String
    Dim rLo      As Long, rHi As Long
    Dim sName    As String
    Dim nFound   As Long
    Dim attrs    As Long

    ' Xay dung pattern: "C:\Windows\*.exe"
    If Right$(FolderPath, 1) = "\" Then
        sPattern = FolderPath & Ext
    Else
        sPattern = FolderPath & "\" & Ext
    End If

    ' Cap phat buffer WIN32_FIND_DATAW
    bLo = X64BufAllocLo(WIN32_FIND_DATA_SIZE)
    bHi = X64BufAllocHi(WIN32_FIND_DATA_SIZE)
    If bLo = 0 And bHi = 0 Then
        FindFiles64 = -1
        Exit Function
    End If

    nFound = 0
    ReDim results(0)

    ' --- FindFirstFileW("C:\Windows\*.exe", &findData) ---
    ' Tham so:
    '   arg0: "W:C:\Windows\*.exe"  (Unicode string, tu dong cap phat buffer)
    '   arg1: pointer den WIN32_FIND_DATAW buffer (truyen Lo/Hi truc tiep)
    '
    ' Call64Ex giu buffer "W:" lai -> nguoi goi khong can doc
    ' Nhung o day can truyen bLo/bHi co san -> dung Call64 + truyen buffer thu cong
    '
    ' Cach: truyen buffer bLo/bHi nhu so nguyen (Lo va Hi rieng)
    ' nArgs=2: arg0=pattern(W:), arg1=bLo (bHi=0 vi WOW64 <4GB)

    Dim p(1) As X64Param_   ' dung type noi bo de khong xung dot
    ' Khong the dung X64Param truc tiep vi la Private Type
    ' -> dung Call64 voi tham so mix: W: cho string, Long cho buffer
    '
    ' Giai phap: truyen buffer pointer qua CLng(bLo) vi bHi=0 trong WOW64

    r = Call64("kernel32.dll", "FindFirstFileW", _
               "W:" & sPattern, CLng(bLo))

    If Not IsRetOK(r) Then
        X64BufFree bLo, bHi
        FindFiles64 = -1
        Exit Function
    End If

    ' FindFirstFileW tra ve HANDLE (32-bit trong WOW64)
    hFind = ParseRetLo(r)
    If hFind = INVALID_HANDLE Or hFind = 0 Then
        X64BufFree bLo, bHi
        FindFiles64 = 0
        Exit Function
    End If

    ' --- Duyet danh sach ---
    Do
        ' Doc ten file tu buffer
        sName = ReadFileName(bLo, bHi)

        ' Bo qua "." va ".."
        If sName <> "." And sName <> ".." Then
            ' Kiem tra khong phai thu muc (neu muon chi lay file)
            attrs = X64BufReadDW(bLo, bHi, OFS_ATTRIBUTES)
            If (attrs And FILE_ATTR_DIRECTORY) = 0 Then
                ' Them vao mang ket qua
                If nFound > 0 Then ReDim Preserve results(nFound)
                results(nFound) = sName
                nFound = nFound + 1
            End If
        End If

        ' FindNextFileW(hFind, &findData)
        ' hFind la HANDLE 32-bit -> truyen nhu Long
        r = Call64("kernel32.dll", "FindNextFileW", _
                   hFind, CLng(bLo))

        If Not IsRetOK(r) Then Exit Do
        If ParseRetLo(r) = 0 Then Exit Do  ' FALSE = het file

    Loop

    ' FindClose(hFind)
    Call64 "kernel32.dll", "FindClose", hFind

    X64BufFree bLo, bHi

    If nFound = 0 Then ReDim results(0)
    FindFiles64 = nFound
End Function



' ============================================================
' Test chinh: tim file trong thu muc, in ra Debug.Print
' ============================================================
Public Sub Test_FindFiles64(Optional ByVal FolderPath As String = "C:\Windows", _
                             Optional ByVal Ext As String = "*.exe")
    Dim Files() As String
    Dim n       As Long
    Dim i       As Long
    Dim msg     As String

    n = FindFiles64(FolderPath, Ext, Files)

    If n < 0 Then
        MsgBox "FindFiles64 that bai (loi khoi tao hoac buffer)", _
               vbCritical, "FindFiles64"
        Exit Sub
    End If

    If n = 0 Then
        MsgBox "Khong tim thay file nao trong:" & vbCrLf & _
               FolderPath & "\" & Ext, vbInformation, "FindFiles64"
        Exit Sub
    End If

    msg = "Tim thay " & n & " file trong " & FolderPath & "\" & Ext & ":" & vbCrLf
    For i = 0 To n - 1
        Debug.Print Files(i)
        If i < 20 Then   ' chi hien 20 file dau trong MsgBox
            msg = msg & "  " & Files(i) & vbCrLf
        ElseIf i = 20 Then
            msg = msg & "  ... (xem them trong Immediate Window)" & vbCrLf
        End If
    Next i

    MsgBox msg, vbInformation, "FindFiles64: " & n & " file"
    Call64_Free
End Sub

' ============================================================
' Wrapper ngan gon: tim va tra ve danh sach ten file
' Su dung: FileList = FindFileList64("C:\Windows", "*.dll")
' ============================================================
Public Function FindFileList64(ByVal FolderPath As String, _
                                ByVal Ext As String) As String
    Dim Files() As String
    Dim n       As Long
    Dim i       As Long
    Dim sOut    As String

    n = FindFiles64(FolderPath, Ext, Files)

    If n <= 0 Then
        FindFileList64 = ""
        Exit Function
    End If

    sOut = ""
    For i = 0 To n - 1
        If Len(sOut) > 0 Then sOut = sOut & vbCrLf
        sOut = sOut & Files(i)
    Next i
    FindFileList64 = sOut
End Function


' Test nhanh
Rem Test_FindFiles64 "C:\Windows", "*.exe"
Sub Test_Call64_FindFiles64()
    ' Hoac goi truc tiep
    Dim Files() As String
    Dim n As Long
    Dim i As Long
    
    n = FindFiles64("C:\Windows\System32", "*.dll", Files)
    For i = 0 To n - 1
        Debug.Print Files(i)
    Next i
End Sub

