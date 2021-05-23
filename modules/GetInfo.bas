Attribute VB_Name = "GetInfo"
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetVersionEx Lib "KERNEL32" Alias _
       "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function getComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Sub GetSystemInfo Lib "KERNEL32" (lpSystemInfo As SYSTEM_INFO)

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 '  Maintenance string for PSS usage
End Type

' dwPlatforID Constants
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

'processor type
Private Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

Public Enum etProcessorType
    PROCESSOR_INTEL_386 = 386
    PROCESSOR_INTEL_486 = 486
    PROCESSOR_INTEL_PENTIUM = 586
    PROCESSOR_MIPS_R4000 = 4000
    PROCESSOR_ALPHA_21064 = 21064
End Enum

Private m_typSystemInfo As SYSTEM_INFO

Private Declare Function GetEnvironmentVariable Lib "KERNEL32" Alias _
  "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function GetDomainName() As String

    Dim lpBuffer As String
    Dim nSize As Long
    Dim lngRetVal As Long
    
    lpBuffer = Space(255)
    nSize = 254
    lngRetVal = GetEnvironmentVariable("USERDOMAIN", lpBuffer, nSize)
    
    GetDomainName = "Domain Name = " + StripNullTerminator(lpBuffer) + vbCrLf
    
End Function

Public Function StripNullTerminator(lpBuffer As String) As String

    Dim i As Integer

    For i = 1 To 255
        If Asc(Mid(lpBuffer, i, 1)) = 0 Then
            lpBuffer = Left(lpBuffer, i - 1)
            Exit For
        End If
    Next i
    
    StripNullTerminator = lpBuffer

End Function

Public Function NumberOfProcessors() As Long
    GetSystemInfo m_typSystemInfo
    NumberOfProcessors = m_typSystemInfo.dwNumberOfProcessors
End Function

Public Function myUserName()

  Dim m_myBuf As String * 25
  Dim m_Val As Long, UserName As String

  m_Val = GetUserName(m_myBuf, 25)
  UserName = Left(m_myBuf, InStr(m_myBuf, Chr(0)) - 1)
  frmMem.lblUserName.Caption = "User Name: " & UserName

End Function

Public Function ComputerName()
  Dim compName As String
  Dim retVal As Long
  
  compName = Space(255)
  retVal = getComputerName(compName, 255)
  If retVal = 0 Then
     MsgBox "Could not get computer name."
     Exit Function
  End If
  'Remove the null character from the end
  compName = Left(compName, InStr(compName, vbNullChar) - 1)
  frmMem.lblComputerName.Caption = "Computer Name: " & compName
 
End Function

Public Function ProcessorType() As etProcessorType
    'See declarations for meaning of returned values
    GetSystemInfo m_typSystemInfo
    ProcessorType = m_typSystemInfo.dwProcessorType
End Function

