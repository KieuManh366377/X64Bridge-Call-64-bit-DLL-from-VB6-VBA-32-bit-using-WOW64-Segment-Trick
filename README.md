# 🌉 X64Bridge.dll

> **Cầu nối miễn phí** — cho phép ứng dụng **32-bit** (VB6, VBA, Delphi Win32) gọi trực tiếp hàm trong bất kỳ **DLL 64-bit** nào trên Windows, không cần viết lại ứng dụng.

---

## 🙏 Lời tri ân

Dự án này được xây dựng dựa trên kỹ thuật **WOW64 Segment Trick** được khám phá và công bố bởi **[The trick (thetrik)](https://github.com/thetrik/Vb64BitDllUsage)**.

> *"Theo tôi, thetrik là một hacker siêu hạng — người đã một tay khai phá ra phương pháp này, mở ra cánh cửa mà nhiều người tưởng là không thể."*

Toàn bộ ý tưởng cốt lõi — JMP FAR `0x33` vào long mode, đọc PEB64 qua `NtWow64ReadVirtualMemory64`, parse export table PE64 từ WOW64 process — đến từ nghiên cứu tiên phong của ông.

🔗 **Nguồn gốc kỹ thuật:** https://github.com/thetrik/Vb64BitDllUsage

---

## 📦 Tải về

| File | Mô tả |
|------|-------|
| `X64Bridge.dll` | DLL 32-bit (Win32), dùng cho VB6 / VBA 32-bit / Delphi Win32 |

> ⚠️ **Lưu ý:** DLL này **phải là 32-bit**. Chỉ hoạt động khi ứng dụng gọi là **32-bit chạy trên Windows 64-bit** (WOW64).

---

## 🧠 Nguyên lý hoạt động

```
Ứng dụng 32-bit (VB6 / VBA / Delphi Win32)
          │
          │  gọi X64Bridge.dll (Win32)
          ▼
┌─────────────────────────────────────────────┐
│         X64Bridge.dll (Win32)               │
│                                             │
│  1. Đọc PEB64 qua NtWow64ReadVirtualMemory  │
│  2. Walk LDR → tìm base DLL 64-bit         │
│  3. Parse PE64 export table                 │
│  4. Sinh x64 machine code động:            │
│       JMP FAR 0x33:entry64  →  long mode   │
│       MOV RCX/RDX/R8/R9, args             │
│       CALL RAX                             │
│       JMP FAR 0x23:ret32   →  32-bit mode  │
└─────────────────────────────────────────────┘
          │
          │  gọi hàm 64-bit thực sự
          ▼
    DLL 64-bit bất kỳ
  (ntdll, kernel32, DLL Go, ...)
```

Điều kỳ diệu: trong WOW64, CPU có thể chuyển đổi giữa **32-bit** (`CS=0x23`) và **64-bit** (`CS=0x33`) chỉ bằng một lệnh `JMP FAR`. Đây chính là bí mật mà thetrik đã khám phá.

---

## 🔌 API xuất ra

```
X64Init()           — Khởi tạo (gọi 1 lần trước khi dùng)
X64Free()           — Giải phóng tài nguyên

X64Invoke()         — Gọi hàm 64-bit (ANSI — VB6)
X64InvokeW()        — Gọi hàm 64-bit (Unicode — VBA/Delphi)

X64BufAllocLo/Hi()  — Cấp phát buffer để truyền pointer
X64BufFree()        — Giải phóng buffer
X64BufWriteAnsi()   — Ghi ANSI string vào buffer
X64BufWriteUnicode()— Ghi Unicode string vào buffer
X64BufReadDW()      — Đọc DWORD từ buffer

X64ModLoA/HiA()     — GetModuleHandle64 (ANSI, cho VB6)
X64Version()        — Phiên bản DLL
```

---

## 💻 Cách dùng — VB6

### Khai báo

```vb
' Kieu du lieu
Private Type X64Param
    Lo As Long
    Hi As Long
End Type

' API chinh
Private Declare Function X64Init   Lib "X64Bridge.dll" () As Long
Private Declare Sub      X64Free   Lib "X64Bridge.dll" ()

Private Declare Function X64Invoke Lib "X64Bridge.dll" ( _
    ByVal sDll      As String,   _
    ByVal sFuncName As String,   _
    ByVal nArgs     As Long,     _
    ByRef  pParams  As X64Param, _
    ByRef  rLo      As Long,     _
    ByRef  rHi      As Long) As Long

Private Declare Function X64BufAllocLo Lib "X64Bridge.dll" (ByVal n As Long) As Long
Private Declare Function X64BufAllocHi Lib "X64Bridge.dll" (ByVal n As Long) As Long
Private Declare Sub      X64BufFree    Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long)
Private Declare Function X64BufReadDW  Lib "X64Bridge.dll" (ByVal bLo As Long, ByVal bHi As Long, ByVal ofs As Long) As Long
```

### Ví dụ 1 — Gọi hàm không tham số

```vb
' NtGetCurrentProcessorNumber() — lay so CPU dang chay
Dim p(0) As X64Param
Dim rLo As Long, rHi As Long

X64Init
X64Invoke "ntdll.dll", "NtGetCurrentProcessorNumber", 0, p(0), rLo, rHi
MsgBox "CPU index = " & rLo
X64Free
```

### Ví dụ 2 — RtlGetVersion (struct output)

```vb
' Lay phien ban Windows chinh xac, khong bi compatibility shim
Dim p(0) As X64Param
Dim rLo As Long, rHi As Long
Dim bLo As Long, bHi As Long

X64Init
bLo = X64BufAllocLo(148)    ' RTL_OSVERSIONINFOW = 148 byte
bHi = X64BufAllocHi(148)

' Ghi dwOSVersionInfoSize = 148 tai offset 0
Dim hdr As String * 4
Mid$(hdr,1,1) = Chr$(148) : Mid$(hdr,2,1) = Chr$(0)
Mid$(hdr,3,1) = Chr$(0)   : Mid$(hdr,4,1) = Chr$(0)
X64BufWriteAnsi bLo, bHi, hdr

p(0).Lo = bLo : p(0).Hi = bHi
X64Invoke "ntdll.dll", "RtlGetVersion", 1, p(0), rLo, rHi

MsgBox "Windows " & X64BufReadDW(bLo, bHi, 4) & _
       "."        & X64BufReadDW(bLo, bHi, 8) & _
       " Build "  & X64BufReadDW(bLo, bHi, 12)

X64BufFree bLo, bHi
X64Free
```

### Ví dụ 3 — GetTickCount64

```vb
Dim p(0) As X64Param
Dim rLo As Long, rHi As Long

X64Init
X64Invoke "kernel32.dll", "GetTickCount64", 0, p(0), rLo, rHi

Dim ms As Currency
ms = CCur(rLo And &H7FFFFFFF)
If (rLo And &H80000000) <> 0 Then ms = ms + CCur(2147483648#)

Dim sec As Long: sec = CLng(ms / 1000)
MsgBox "Uptime: " & sec\3600 & "h " & (sec Mod 3600)\60 & "m " & sec Mod 60 & "s"
X64Free
```

---

## 💻 Cách dùng — VBA (Excel/Access 32-bit)

```vb
' VBA dung PtrSafe nhung KHONG dung PtrSafe voi X64Bridge (DLL 32-bit)
' Chi chay trong host 32-bit: Excel 32-bit, Access 32-bit, VB6

Private Declare Function X64Init Lib "X64Bridge.dll" () As Long
' ... tuong tu VB6
```

> ⚠️ **Không dùng được với Excel 64-bit** — vì Excel 64-bit không thể load DLL 32-bit.

---

## ✅ Kết quả kiểm thử

Kiểm thử trên **Windows 11 Build 26200**, VB6 SP6:

| Test | Kết quả |
|------|---------|
| `NtGetCurrentProcessorNumber` (0 args) | ✅ PASS |
| `NtQuerySystemTime` (1 output ptr) | ✅ PASS |
| `RtlGetVersion` (struct in/out) | ✅ PASS |
| `GetTickCount64` (64-bit result) | ✅ PASS |
| `NtQueryPerformanceCounter` (2 args) | ✅ PASS |
| `NtQueryInformationProcess` (5 args) | ✅ PASS |
| `FindFirstFileW` / `FindNextFileW` | ✅ PASS |
| System32 64-bit (3527 DLL) | ✅ Thấy đủ (Dir() chỉ thấy 2348) |

> **Lý do thấy nhiều file hơn `Dir()`:** Windows tự redirect `C:\Windows\System32` → `SysWOW64` với process 32-bit. `X64Bridge` gọi `FindFirstFileW` 64-bit thật nên thấy `System32` thực sự.

---

## ⚙️ Yêu cầu hệ thống

| Yêu cầu | Chi tiết |
|---------|---------|
| OS | Windows 7/8/10/11 **64-bit** |
| App gọi | **32-bit** (VB6, VBA 32-bit, Delphi Win32) |
| Runtime | Không cần cài thêm gì |
| Quyền | User thường (không cần Admin) |

---

## 📋 Ghi chú kỹ thuật

- **Tham số âm** (`-1`, handle process...): cần truyền đúng dạng sign-extended 64-bit — `Lo=&HFFFFFFFF, Hi=&HFFFFFFFF`
- **Buffer pointer**: `VirtualAlloc` trong WOW64 luôn cấp phát dưới 4GB → `Hi` của buffer luôn = 0
- **Tên hàm export**: PE export table luôn dùng ANSI — truyền `ByVal String` trong VB6 là đúng
- **Thread-safe**: mỗi lần gọi `X64Invoke` sinh lại machine code vào code buffer → **không thread-safe**, gọi từ một thread duy nhất

---

## 📜 License

**Miễn phí sử dụng** cho mọi mục đích — cá nhân, thương mại, học thuật.

Chỉ chia sẻ file DLL đã biên dịch. Mã nguồn không được công bố.

---

## 👤 Tác giả

**Kieu Manh**
📧 kieumanh366377@gmail.com

---

## 🔗 Tham khảo

- **Kỹ thuật gốc (thetrik):** https://github.com/thetrik/Vb64BitDllUsage
- **WOW64 internals:** https://docs.microsoft.com/en-us/windows/win32/winprog64/wow64-implementation-details
- **PE format:** https://docs.microsoft.com/en-us/windows/win32/debug/pe-format

---

<div align="center">

*Được xây dựng với ❤️ trên Delphi 13.1 & C++ Builder 13.1*

*Kỹ thuật WOW64 Segment Trick © thetrik — mọi công nhận xứng đáng thuộc về ông*

</div>
