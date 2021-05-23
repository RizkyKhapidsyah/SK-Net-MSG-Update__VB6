Attribute VB_Name = "System"
Option Explicit

Public szUserInfo As String
Public hKey As Long
Public Success As Boolean
Public InMegs As Boolean

Public Const szSubkey = "SOFTWARE\Microsoft\Windows\CurrentVersion"

Public Const AB_NO_USER = &H1
Public Const AB_NO_COMPANY = &H2
Public Const AB_NO_WIN_VERSION = &H4
Public Const AB_NO_VERSION_NUMBER = &H8
Public Const AB_NO_BUILD_NUMBER = &H10
Public Const AB_NO_CPU = &H20
Public Const AB_NO_PHYSICAL = &H40
Public Const AB_NO_PAGING = &H80
Public Const AB_NO_VIRTUAL = &H100
Public Const AB_NO_MEMLOAD = &H200

' O/S Version Info structure
' Used to get the operating system version and platform information
Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type
' dwPlatformId defines for OSVERSIONINFO structure...
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2
' and related Win API call...
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

' System Info structure
' Used to get the amount and type of CPU information
Type SYSTEM_INFO
    dwOemID                     As Long
    dwPageSize                  As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask       As Long
    dwNumberOfProcessors        As Long
    dwProcessorType             As Long
    dwAllocationGranularity     As Long
    dwReserved                  As Long
End Type
' dwProcessorType defines for SYSTEM_INFO structure...
Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_PENTIUM = 586
Public Const PROCESSOR_MIPS_R2000 = 2000
Public Const PROCESSOR_MIPS_R3000 = 3000
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064
' and related Win API call...
Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

' Memory Status Info structure
' Used to get various system memory information
Type MEMORYSTATUS
    dwLength        As Long  ' sizeof(MEMORYSTATUS)
    dwMemoryLoad    As Long  ' percent of memory in use (between 1 and 100)
    dwTotalPhys     As Long  ' bytes of physical memory
    dwAvailPhys     As Long  ' free physical memory bytes
    dwTotalPageFile As Long  ' bytes of paging file
    dwAvailPageFile As Long  ' free bytes of paging file
    dwTotalVirtual  As Long  ' user bytes of address space
    dwAvailVirtual  As Long  ' free user bytes
End Type
' and related Win API call...
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

' Registry manipulation API's for getting the User or Company name
Declare Function OSRegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulSSOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function OSRegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Declare Function OSRegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, lpdwType As Long, lpbData As Any, cbData As Long) As Long

Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Public Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_QUERY_VALUE = &H1
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1

Public MemStat As MEMORYSTATUS
Public MemData As Long

Public Function MSysInfo()

    Dim tOSVer As OSVERSIONINFO
   ' First set length of OSVERSIONINFO structure size
   tOSVer.dwOSVersionInfoSize = Len(tOSVer)
   ' Get version information
   GetVersionEx tOSVer
   ' Determine OS type
   With tOSVer
      
      Select Case .dwPlatformId
         Case VER_PLATFORM_WIN32_NT
            ' This is an NT version (NT/2000)
            ' If dwMajorVersion >= 5 then the OS is Win2000
            If .dwMajorVersion >= 5 Then
               SysInfo.Label1.Caption = "Windows 2000"
            Else
               SysInfo.Label1.Caption = "Windows NT"
            End If
         Case Else
            ' This is Windows 95/98/ME
            If .dwMajorVersion >= 5 Then
               SysInfo.Label1.Caption = "Windows ME"
            ElseIf .dwMajorVersion = 4 And .dwMinorVersion > 0 Then
               SysInfo.Label1.Caption = "Windows 98"
            Else
               SysInfo.Label1.Caption = "Windows 95"
            End If
         End Select
         ' Check for service pack
         SysInfo.Label1.Caption = SysInfo.Label1.Caption & " " & Left(.szCSDVersion, _
                          InStr(1, .szCSDVersion, Chr$(0)))
         ' Get OS version
         SysInfo.Label2.Caption = "Version: " & .dwMajorVersion & "." & _
                          .dwMinorVersion & "." & .dwBuildNumber
        
    End With
    
End Function
Function GetBuildNumber() As String
  Dim OSVer As OSVERSIONINFO
  Dim lResult As Long
  
  OSVer.dwOSVersionInfoSize = Len(OSVer)
  lResult = GetVersionEx(OSVer)
  GetBuildNumber$ = "Build:  " & Format$(OSVer.dwBuildNumber Mod 65536)
End Function

Function GetCompany() As String
  If (OSRegOpenKeyEx(HKEY_LOCAL_MACHINE, szSubkey, 0&, KEY_QUERY_VALUE, hKey)) = ERROR_SUCCESS Then
    Success = RegQueryStringValue(hKey, "RegisteredOrganization", szUserInfo)
    Success = RegCloseKey(hKey)
    GetCompany$ = szUserInfo
  Else
    GetCompany$ = "Not listed..."
  End If
End Function

Function GetCPUType() As String
  Dim SysInfo As SYSTEM_INFO
  Dim CPU_Name As String
  Call GetSystemInfo(SysInfo)
  Select Case SysInfo.dwProcessorType
  Case PROCESSOR_INTEL_386
    CPU_Name = "Intel 386"
  Case PROCESSOR_INTEL_486
    CPU_Name = "Intel 486"
  Case PROCESSOR_INTEL_PENTIUM
    CPU_Name = "Pentium"
  Case PROCESSOR_MIPS_R2000
    CPU_Name = "Mips R2000"
  Case PROCESSOR_MIPS_R3000
    CPU_Name = "Mips R3000"
  Case PROCESSOR_MIPS_R4000
    CPU_Name = "Mips R4000"
  Case PROCESSOR_ALPHA_21064
    CPU_Name = "Alpha 21064"
  Case Else ' default if not defined...
    CPU_Name = Format$(SysInfo.dwProcessorType)
  End Select
  GetCPUType$ = Format$(SysInfo.dwNumberOfProcessors) _
                         & "  " & CPU_Name & "  Processor"
End Function


Function GetMemoryLoad() As String
  MemStat.dwLength = Len(MemStat)
  Call GlobalMemoryStatus(MemStat)
  MemData = MemStat.dwMemoryLoad
  GetMemoryLoad$ = Format$(MemData)
End Function

Function GetOS() As String
  Dim OSVer As OSVERSIONINFO
  Dim lResult As Long
  
  OSVer.dwOSVersionInfoSize = Len(OSVer)
  lResult = GetVersionEx(OSVer)
  If lResult Then
    Select Case OSVer.dwPlatformId
    Case VER_PLATFORM_WIN32s
      GetOS$ = "Win32s Subsystem on Windows 3.xx"
    Case VER_PLATFORM_WIN32_WINDOWS
      GetOS$ = "Microsoft Windows 95"
    Case VER_PLATFORM_WIN32_NT
      GetOS$ = "Microsoft Windows NT"
    End Select
  End If
End Function
Function GetPagingMemory(InKb As Long) As String
  MemStat.dwLength = Len(MemStat)
  Call GlobalMemoryStatus(MemStat)
  MemData = MemStat.dwTotalPageFile
  If Not InKb Then
    GetPagingMemory$ = Format$(MemData)
  Else
    GetPagingMemory$ = Format$(MemData \ 1024, "###,###,###")
  End If
End Function

Function GetFreePagingMemory(InKb As Long) As String
  MemStat.dwLength = Len(MemStat)
  Call GlobalMemoryStatus(MemStat)
  MemData = MemStat.dwAvailPageFile
  If Not InKb Then
    GetFreePagingMemory$ = Format$(MemData)
  Else
    GetFreePagingMemory$ = Format$(MemData \ 1024, "###,###,###")
  End If
End Function

Function GetFreePhysicalMemory(InKb As Long) As String
  MemStat.dwLength = Len(MemStat)
  Call GlobalMemoryStatus(MemStat)
  MemData = MemStat.dwAvailPhys
  If Not InKb Then
    GetFreePhysicalMemory$ = Format$(MemData)
  Else
    GetFreePhysicalMemory$ = Format$(MemData \ 1024, "###,###,###")
  End If
End Function

Function GetPhysicalMemory(InKb As Long) As String
  MemStat.dwLength = Len(MemStat)
  Call GlobalMemoryStatus(MemStat)
  MemData = MemStat.dwTotalPhys
  If Not InKb Then
    GetPhysicalMemory$ = Format$(MemData)
  Else
    GetPhysicalMemory$ = Format$(MemData \ 1024, "###,###,###")
  End If
End Function

Function GetUserName() As String
  If (OSRegOpenKeyEx(HKEY_LOCAL_MACHINE, szSubkey, 0&, KEY_QUERY_VALUE, hKey)) = ERROR_SUCCESS Then
    Success = RegQueryStringValue(hKey, "RegisteredOwner", szUserInfo)
    Success = RegCloseKey(hKey)
    GetUserName$ = szUserInfo
  Else
    GetUserName$ = "Not on network..."
  End If
End Function


Function GetFreeVirtualMemory(InKb As Long) As String
  MemStat.dwLength = Len(MemStat)
  Call GlobalMemoryStatus(MemStat)
  MemData = MemStat.dwAvailVirtual
  If Not InKb Then
    GetFreeVirtualMemory$ = Format$(MemData)
  Else
    GetFreeVirtualMemory$ = Format$(MemData \ 1024, "###,###,###")
  End If
End Function

Function GetVirtualMemory(InKb As Long) As String
  MemStat.dwLength = Len(MemStat)
  Call GlobalMemoryStatus(MemStat)
  MemData = MemStat.dwTotalVirtual
  If Not InKb Then
    GetVirtualMemory$ = Format$(MemData)
  Else
    GetVirtualMemory$ = Format$(MemData \ 1024, "###,###,###")
  End If
End Function

Function GetWinVersion() As String
  Dim OSVer As OSVERSIONINFO
  Dim lResult As Long
  
  OSVer.dwOSVersionInfoSize = Len(OSVer)
  lResult = GetVersionEx(OSVer)
  GetWinVersion$ = "Version:  " & Format$(OSVer.dwMajorVersion) _
                        & "." & Format$(OSVer.dwMinorVersion, "00")
End Function


