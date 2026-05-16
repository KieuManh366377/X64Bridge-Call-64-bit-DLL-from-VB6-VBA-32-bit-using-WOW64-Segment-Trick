VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ============================================================
' Khai bao API
' ============================================================
Private Declare Function SetDllDirectoryW Lib "kernel32" (ByVal lpPath As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleFileNameA Lib "kernel32" (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
' ============================================================
' Kiem tra DLL dang load tu dau
' ============================================================
Public Sub CheckDllPath()
    Dim buf  As String
    Dim hMod As Long
    Dim nLen As Long
    hMod = GetModuleHandleA("X64Bridge.dll")

    If hMod = 0 Then
        Debug.Print "X64Bridge.dll chua duoc load vao process!"
        Debug.Print "Kiem tra lai thu muc va ten file."
        Exit Sub
    End If
    buf = Space(512)
    nLen = GetModuleFileNameA(hMod, buf, 512)

    If nLen > 0 Then
        Debug.Print "DLL dang load tu: " & Left(buf, nLen)
    Else
        Debug.Print "GetModuleFileNameA that bai (hMod=" & Hex(hMod) & ")"
    End If
End Sub

' ============================================================
' Form_Load
' ============================================================
Private Sub Form_Load()
    Dim rc As Long
    ' Buoc 1: Chi ra thu muc chua X64Bridge.dll
    ' App.Path = thu muc chua file .vbp (khi chay trong IDE)
    ' = thu muc chua file .exe (khi chay ngoai IDE)
    rc = SetDllDirectoryW(StrPtr(App.path))
    ' Debug.Print "SetDllDirectoryW('" & App.Path & "'): " & rc
    ' Buoc 2: Load DLL tuong minh truoc khi goi bat ky ham nao
    ' LoadLibrary dam bao Windows tim DLL dung thu muc vua set
    ' rc = LoadLibraryA(App.Path & "\X64Bridge.dll")
    ' If rc = 0 Then
    ' Debug.Print "LoadLibraryA THAT BAI — kiem tra file co ton tai khong:"
    ' Debug.Print "  " & App.Path & "\X64Bridge.dll"
    ' Else
    ' Debug.Print "LoadLibraryA OK (hMod=" & Hex(rc) & ")"
    ' End If
    ' Buoc 3: Xac nhan DLL dang load tu dau
    ' CheckDllPath
    ' Buoc 4: Chay test chinh
    Rem ==========================
     RunAllSteps
    ' Demo_Call64
    ' Demo_Call64
    'Test_Call64_FindFiles64
    ' RunAllSteps
    ' RunAll2
End Sub
