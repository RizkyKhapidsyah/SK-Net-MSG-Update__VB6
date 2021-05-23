VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetMSG v1.0"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMacAdr 
      Caption         =   "Mac Address"
      Height          =   3840
      Left            =   0
      TabIndex        =   21
      Top             =   8130
      Visible         =   0   'False
      Width           =   7965
      Begin VB.CommandButton OKButton 
         Caption         =   "OK"
         Height          =   375
         Left            =   6675
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1125
      End
      Begin VB.TextBox txtMac 
         Height          =   3450
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   315
         Width           =   6405
      End
      Begin VB.CommandButton cmdGetMac 
         Caption         =   "Inspect"
         Height          =   300
         Left            =   6675
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   300
         Width           =   1125
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   300
         Left            =   6675
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   675
         Width           =   1125
      End
   End
   Begin VB.Frame fraMain 
      Caption         =   "NetMSG Plus"
      Height          =   3840
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7965
      Begin VB.CommandButton cmdMin 
         Caption         =   "Minimise"
         Height          =   360
         Left            =   6705
         TabIndex        =   28
         Top             =   1005
         Width           =   1155
      End
      Begin VB.PictureBox PicHook 
         Height          =   465
         Left            =   300
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   27
         ToolTipText     =   "NetMSG v1.0"
         Top             =   2400
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Timer TimerTime 
         Interval        =   1000
         Left            =   300
         Top             =   3000
      End
      Begin VB.CommandButton cmdUser 
         Caption         =   "SysInfo"
         Height          =   360
         Left            =   6705
         TabIndex        =   26
         Top             =   2580
         Width           =   1155
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Lock Station"
         Height          =   360
         Left            =   6705
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "This Locks the Workstation you are on."
         Top             =   1455
         Width           =   1155
      End
      Begin VB.OptionButton optUsers 
         Caption         =   "Users"
         Height          =   255
         Left            =   5805
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "This option requires a user name to be input."
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optDomain 
         Caption         =   "Domain"
         Height          =   255
         Left            =   4725
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "This option requires a valid domain. If left blank it will assume the current domain."
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   360
         Left            =   6705
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "About the Program."
         Top             =   225
         Width           =   1155
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   360
         Left            =   6705
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Uh!"
         Top             =   615
         Width           =   1155
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   360
         Left            =   6705
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Resets the Interface"
         Top             =   3015
         Width           =   1155
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Enabled         =   0   'False
         Height          =   360
         Left            =   6705
         TabIndex        =   12
         ToolTipText     =   "Sends the current message."
         Top             =   3405
         Width           =   1155
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3045
         Left            =   1050
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   705
         Width           =   5475
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1065
         TabIndex        =   10
         Top             =   360
         Width           =   3480
      End
      Begin VB.CommandButton cmdMacAdr 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mac Address"
         Height          =   360
         Left            =   6705
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "This option uses Arp to extract Mac Addresses"
         Top             =   1830
         Width           =   1155
      End
      Begin VB.CommandButton cmdPing 
         Caption         =   "Ping Host"
         Height          =   360
         Left            =   6705
         TabIndex        =   8
         ToolTipText     =   "This option allows you to ping any valid Host / IP Address"
         Top             =   2205
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Message:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   225
         TabIndex        =   20
         Top             =   705
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Domain:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   315
         TabIndex        =   19
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame fraPing 
      Caption         =   "Ping"
      Height          =   3840
      Left            =   0
      TabIndex        =   1
      Top             =   4230
      Visible         =   0   'False
      Width           =   7965
      Begin VB.CommandButton cmdPReturn 
         Caption         =   "OK"
         Height          =   390
         Left            =   6675
         TabIndex        =   6
         Top             =   675
         Width           =   1140
      End
      Begin VB.TextBox txtPingResults 
         Height          =   3090
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   675
         Width           =   6390
      End
      Begin VB.CommandButton cmdPingIt 
         Caption         =   "Ping!"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6675
         TabIndex        =   4
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox txtPingAdr 
         Height          =   315
         Left            =   1575
         TabIndex        =   2
         Text            =   "Enter Host / IP Address to Ping..."
         Top             =   285
         Width           =   4965
      End
      Begin VB.Label lblPing 
         AutoSize        =   -1  'True
         Caption         =   "Ping Destination:"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   375
         Width           =   1200
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   225
      Top             =   1650
   End
   Begin ComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   3840
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   529
      SimpleText      =   ""
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12681
            MinWidth        =   457
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore Window"
      End
      Begin VB.Menu mnuPopUpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSysStats 
         Caption         =   "System Statistics"
      End
      Begin VB.Menu mnuPopUpSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit NetMSG Plus"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NERR_Success As Long = 0&
Private Declare Function LockWorkStation Lib "user32.dll" () As Long

Private Declare Function NetMessageBufferSend Lib "NETAPI32.DLL" _
(yServer As Any, yToName As Byte, yFromName As Any, yMsg As Byte, ByVal lSize As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim Str As String

'###########################################################################################################
' Timer Functions
'###########################################################################################################

Private Sub TimerTime_Timer()
    sb.Panels(2).Text = Now
End Sub

Private Sub Timer1_Timer()
    If Not Text2.Text = "" Then
        cmdSend.Enabled = True
    Else
        cmdSend.Enabled = False
    End If
End Sub

'###########################################################################################################
' Form Functions
'###########################################################################################################

Private Sub Form_Activate()
    'reset to base values
    cmdReset_Click
    
End Sub

Private Sub Form_Load()
    Set Tray_Icon = frmMain.Icon
    Call TRAY_Create(PicHook)

    sb.Panels(2).Text = Now
    TimerTime.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call TRAY_Close(PicHook)
End Sub

Private Sub Form_Paint()
    
    Me.AutoRedraw = True
    Me.Refresh
    
End Sub

Private Sub Form_Resize()
    
    If Me.WindowState = vbMinimized Then
        Me.Visible = False
    Else
        Me.Visible = True
    End If
    Call Form_Paint
End Sub

'###########################################################################################################
' SysTray Icon
'###########################################################################################################

Private Sub PicHook_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim Msg As Long
    Dim dReturn
    Msg = x / Screen.TwipsPerPixelX
    
    Select Case Msg
        Case WM_LBUTTONUP        'Single Click
            dReturn = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_LBUTTONDBLCLK    'Double Click
            Me.WindowState = vbNormal
            dReturn = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_RBUTTONUP
            PopupMenu Me.mnuPopUp
    End Select
 
End Sub

'###########################################################################################################
' Systray Menu Items
'###########################################################################################################

Private Sub mnuRestore_Click()
    Load frmMain
    frmMain.WindowState = vbNormal
    frmMain.Show
End Sub
Private Sub mnuSysStats_Click()
    Load frmMem
    frmMem.Show , Me
End Sub

Private Sub mnuExit_Click()
    Unload Me: End
End Sub

'###########################################################################################################
' Lock Workstation
'###########################################################################################################

Private Sub cmdLock_Click()
    Dim bUserCancel As Boolean
    Call LockWorkStation
End Sub

'###########################################################################################################
'Deals with the Ping Command
'###########################################################################################################

Private Sub cmdPing_Click()
    fraPing.Top = 0
    fraPing.Left = 0
    fraMain.Visible = False
    fraPing.Visible = True
    txtPingResults.Text = ""
    txtPingAdr.Text = "Enter Host / IP Address to Ping..."
    sb.Panels(1).Text = "Ping a Host / IP Address"
End Sub

Private Sub cmdPReturn_Click()
    Call txtPingAdr_Click
    fraPing.Visible = False
    fraMain.Visible = True
    sb.Panels(1).Text = App.Title
        cmdPingIt.Enabled = False
End Sub

Private Sub cmdPingIt_Click()
    txtPingResults.Text = ""
    txtPingResults.Text = MGetCmdOutput.GetCommandOutput("ping " & txtPingAdr.Text, True, False, True)
End Sub

Private Sub cmdUser_Click()
    'myUserName
    Load frmMem
    frmMem.Show , Me
    sb.Panels(1).Text = "Viewing Current User and Computer Name."
End Sub

Private Sub txtPingAdr_Click()
    txtPingAdr.Text = "127.0.0.1"
    txtPingResults.Text = ""
    sb.Panels(1).Text = "Pinging: " & txtPingAdr.Text
        cmdPingIt.Enabled = True
End Sub

Private Sub txtPingAdr_Change()
    If Not txtPingAdr.Text = "" Then sb.Panels(1).Text = "Pinging: " & txtPingAdr.Text
End Sub

'###########################################################################################################
' Mac Address Function
'###########################################################################################################

Private Sub cmdMacAdr_Click()
    fraMacAdr.Top = 0
    fraMacAdr.Left = 0
    fraMain.Visible = False
    fraMacAdr.Visible = True
    txtMac.Text = ""
    sb.Panels(1).Text = "Mac Address Extractor"
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText txtMac.Text
End Sub

Private Sub cmdGetMac_Click()
    txtMac.Text = MGetCmdOutput.GetCommandOutput("arp -a", True, False, True)
End Sub

Private Sub OKButton_Click()
    fraMacAdr.Visible = False
    fraMain.Visible = True
    sb.Panels(1).Text = App.Title
End Sub

'###########################################################################################################
' Button Commands
'###########################################################################################################

Private Sub cmdMin_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub cmdAbout_Click()
    sb.Panels(1).Text = "About NetMsg v1.0"
    Load frmAbout
    frmAbout.Show , Me
    sb.Panels(1).Text = App.Title
End Sub

Private Sub cmdExit_Click()
    Unload Me: End
End Sub

Private Sub cmdReset_Click()
    Text1.Text = "/DOMAIN:"
    Text2.Text = ""
    Text2.SetFocus
    Text1.SetFocus
    optUsers.Value = False
    optDomain.Value = True
End Sub

Private Sub cmdSend_Click()
    Dim c, m As String
    c = Text1.Text
    m = Text2.Text
    
    If (Text1.Text = "") Then
        MsgBox "Enter Computer/User Name"
        Text1.SetFocus
    ElseIf (Text2.Text = "") Then
        MsgBox "Enter Your Message"
        Text2.SetFocus
    End If
    
    Call MGetCmdOutput.GetCommandOutput("net send " & c & " " & m, True, False, True)
    'Debug.Print MGetCmdOutput.GetCommandOutput("net send " & c & " " & m, True, False, True)
    
End Sub

'###########################################################################################################
' Option Logic (:?)
'###########################################################################################################

Private Sub Text1_GotFocus()
    sb.Panels(1).Text = "Enter Details"
    Timer1.Enabled = False
End Sub

Private Sub Text2_GotFocus()
    sb.Panels(1).Text = "Enter Your Message"
    Timer1.Enabled = True
End Sub

Private Sub optDomain_Click()
    If optUsers.Value = True Then
        optDomain.Value = False
        Label1.Caption = "User:"
        Text1.Text = "Enter User Name"
    ElseIf optUsers.Value = False Then
        optDomain.Value = True
        Label1.Caption = "Domain:"
        Text1.Text = "/DOMAIN:"
    End If
End Sub

Private Sub optDomain_GotFocus()
    sb.Panels(1).Text = "This option uses the current domain."
End Sub

Private Sub optUsers_Click()
    If optDomain.Value = True Then
        optUsers.Value = False
        Label1.Caption = "Domain:"
        Text1.Text = "/DOMAIN:"
    ElseIf optDomain.Value = False Then
        optUsers.Value = True
        Label1.Caption = "User:"
        Text1.Text = "Enter User Name"
    End If
End Sub

Private Sub optUsers_GotFocus()
    sb.Panels(1).Text = "This option requires a user name."
End Sub


