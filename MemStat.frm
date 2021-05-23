VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMem 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Status"
   ClientHeight    =   4230
   ClientLeft      =   2295
   ClientTop       =   1350
   ClientWidth     =   4695
   Icon            =   "MemStat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4695
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3990
      Picture         =   "MemStat.frx":0CCA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   150
      Width           =   480
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3825
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4005
      Top             =   675
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   4560
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   255
      Left            =   195
      TabIndex        =   21
      Top             =   165
      Width           =   2655
   End
   Begin VB.Label lblComputerName 
      BackStyle       =   0  'Transparent
      Caption         =   "Computer Name:"
      Height          =   255
      Left            =   195
      TabIndex        =   20
      Top             =   405
      Width           =   2655
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   195
      TabIndex        =   19
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   195
      TabIndex        =   18
      Top             =   690
      Width           =   2655
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   240
      Left            =   195
      TabIndex        =   17
      Top             =   1245
      Width           =   3735
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   4560
      Y1              =   3465
      Y2              =   3465
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Used Memory:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3585
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1890
      TabIndex        =   13
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Virtual Bytes:"
      Height          =   255
      Left            =   180
      TabIndex        =   12
      Top             =   3135
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1635
      TabIndex        =   11
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Virtual bytes:"
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   2895
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2355
      TabIndex        =   9
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Available bytes of Page File:"
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   2655
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1515
      TabIndex        =   7
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Page File Bytes:"
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   2415
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2250
      TabIndex        =   5
      Top             =   2160
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available Physical Memory:"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   2175
      Width           =   1920
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   1920
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Physical Memory:"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   1935
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1275
      TabIndex        =   1
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Memory used:"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   1695
      Width           =   1095
   End
End
Attribute VB_Name = "frmMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Sub GlobalMemoryStatus Lib "KERNEL32" (lpBuffer As MEMORYSTATUS)
Private Declare Sub ZeroMemory Lib "KERNEL32" Alias "RtlMoveMemory" (dest As Any, ByVal numBytes As Long)

Dim Mem As MEMORYSTATUS

Private Type MEMORYSTATUS
    dwlength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Type NOTIFYICONDATA
     cbSize As Long
     hwnd As Long
     uID As Long
     uFlags As Long
     uCallbackMessage As Long
     hIcon As Long
     szTip As String * 64
End Type

Const NIM_ADD = &H0&
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4
Const WM_MOUSEMOVE = &H200

Dim NI As NOTIFYICONDATA
Dim result As Long
Dim Response As String
Dim TimeToCheck As Integer
Dim ShowAlert As Boolean
Dim Msgs() As String
Dim nDelay As Long
Dim bErr As Boolean, NotCommand As Boolean
Dim Reply As String, sFrom As String, sSubject As String

Private Sub Quit_Click()
    Unload Me
End Sub

Private Sub Form_Load()

GlobalMemoryStatus Mem
'BytesFree.Caption = Mem.dwMemoryLoad & "% used"
Label2.Caption = Mem.dwMemoryLoad & "%"
Label4.Caption = Mem.dwTotalPhys
Label6.Caption = Mem.dwAvailPhys
Label8.Caption = Mem.dwTotalPageFile
Label10.Caption = Mem.dwAvailPageFile
Label12.Caption = Mem.dwTotalVirtual
Label14.Caption = Mem.dwAvailVirtual
ProgressBar1.Value = Mem.dwMemoryLoad

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
               Label17.Caption = "Windows 2000"
            Else
               Label17.Caption = "Windows NT"
            End If
         Case Else
            ' This is Windows 95/98/ME
            If .dwMajorVersion >= 5 Then
               Label17.Caption = "Windows ME"
            ElseIf .dwMajorVersion = 4 And .dwMinorVersion > 0 Then
               Label17.Caption = "Windows 98"
            Else
               Label17.Caption = "Windows 95"
            End If
         End Select
         ' Check for service pack
         Label17.Caption = Label17.Caption & " " & Left(.szCSDVersion, _
                          InStr(1, .szCSDVersion, Chr$(0)))
         ' Get OS version
         Label18.Caption = "Version: " & .dwMajorVersion & "." & _
                          .dwMinorVersion & "." & .dwBuildNumber
        
    End With
    
    Label16.Caption = "Processor Type: " & ProcessorType & " = " & NumberOfProcessors & " Installed."
    
    Call GetInfo.GetUserName(GetInfo.myUserName, 0)
    Call GetInfo.ComputerName
End Sub

Private Sub Timer1_Timer()
GlobalMemoryStatus Mem
'BytesFree.Caption = Mem.dwMemoryLoad & "% used"
Label2.Caption = Mem.dwMemoryLoad & "%"
Label4.Caption = Mem.dwTotalPhys
Label6.Caption = Mem.dwAvailPhys
Label8.Caption = Mem.dwTotalPageFile
Label10.Caption = Mem.dwAvailPageFile
Label12.Caption = Mem.dwTotalVirtual
Label14.Caption = Mem.dwAvailVirtual
ProgressBar1.Value = Mem.dwMemoryLoad
End Sub



