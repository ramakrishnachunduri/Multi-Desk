VERSION 5.00
Begin VB.Form main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Multi Desk"
   ClientHeight    =   4065
   ClientLeft      =   4815
   ClientTop       =   2985
   ClientWidth     =   9000
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "main.frx":0B3A
   ScaleHeight     =   4065
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtevent 
      Height          =   285
      Left            =   2640
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   3360
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1670
      Left            =   630
      ScaleHeight     =   1665
      ScaleWidth      =   2115
      TabIndex        =   20
      Top             =   1150
      Width           =   2115
   End
   Begin VB.Frame winframe1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Multi Desk 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   7
      Left            =   3120
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
      Begin VB.ListBox winlist1 
         BackColor       =   &H00D8E9EC&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2460
         Index           =   7
         ItemData        =   "main.frx":9A55
         Left            =   240
         List            =   "main.frx":9A57
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame winframe1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Multi Desk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   6
      Left            =   3120
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
      Begin VB.ListBox winlist1 
         BackColor       =   &H00D8E9EC&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2460
         Index           =   6
         ItemData        =   "main.frx":9A59
         Left            =   240
         List            =   "main.frx":9A5B
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame winframe1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Multi Desk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   5
      Left            =   3120
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
      Begin VB.ListBox winlist1 
         BackColor       =   &H00D8E9EC&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2460
         Index           =   5
         ItemData        =   "main.frx":9A5D
         Left            =   240
         List            =   "main.frx":9A5F
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame winframe1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Multi Desk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   4
      Left            =   3120
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
      Begin VB.ListBox winlist1 
         BackColor       =   &H00D8E9EC&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2460
         Index           =   4
         ItemData        =   "main.frx":9A61
         Left            =   240
         List            =   "main.frx":9A63
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame winframe1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Multi Desk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   3
      Left            =   3120
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
      Begin VB.ListBox winlist1 
         BackColor       =   &H00D8E9EC&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2460
         Index           =   3
         ItemData        =   "main.frx":9A65
         Left            =   240
         List            =   "main.frx":9A67
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame winframe1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Multi Desk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   2
      Left            =   3120
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
      Begin VB.ListBox winlist1 
         BackColor       =   &H00D8E9EC&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2460
         Index           =   2
         ItemData        =   "main.frx":9A69
         Left            =   240
         List            =   "main.frx":9A6B
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame winframe1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Multi Desk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   1
      Left            =   3120
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
      Begin VB.ListBox winlist1 
         BackColor       =   &H00D8E9EC&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2460
         Index           =   1
         ItemData        =   "main.frx":9A6D
         Left            =   240
         List            =   "main.frx":9A6F
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame winframe1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Multi Desk 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   0
      Left            =   3120
      TabIndex        =   4
      Top             =   720
      Width           =   3735
      Begin VB.ListBox winlist1 
         BackColor       =   &H00D8E9EC&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2460
         Index           =   0
         ItemData        =   "main.frx":9A71
         Left            =   240
         List            =   "main.frx":9A73
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "    Desktops    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   6840
      TabIndex        =   3
      Top             =   720
      Width           =   1935
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   510
         Left            =   720
         Top             =   2280
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Image NextDesk 
         Height          =   480
         Left            =   1320
         Picture         =   "main.frx":9A75
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image PrevDesk 
         Height          =   480
         Left            =   120
         Picture         =   "main.frx":9EB7
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image btndesk 
         Height          =   480
         Index           =   8
         Left            =   1320
         Picture         =   "main.frx":A2F9
         Top             =   1680
         Width           =   480
      End
      Begin VB.Image btndesk 
         Height          =   480
         Index           =   7
         Left            =   720
         Picture         =   "main.frx":A603
         Top             =   1680
         Width           =   480
      End
      Begin VB.Image btndesk 
         Height          =   480
         Index           =   6
         Left            =   120
         Picture         =   "main.frx":A90D
         Top             =   1680
         Width           =   480
      End
      Begin VB.Image btndesk 
         Height          =   480
         Index           =   5
         Left            =   1320
         Picture         =   "main.frx":AC17
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image btndesk 
         Height          =   480
         Index           =   4
         Left            =   720
         Picture         =   "main.frx":AF21
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image btndesk 
         Height          =   480
         Index           =   3
         Left            =   120
         Picture         =   "main.frx":B22B
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image btndesk 
         Height          =   480
         Index           =   2
         Left            =   1320
         Picture         =   "main.frx":B535
         Top             =   480
         Width           =   480
      End
      Begin VB.Image btndesk 
         Height          =   480
         Index           =   1
         Left            =   720
         Picture         =   "main.frx":B83F
         Top             =   480
         Width           =   480
      End
      Begin VB.Image btndesk 
         Height          =   480
         Index           =   0
         Left            =   120
         Picture         =   "main.frx":BB49
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.Timer hotKeyTimer 
      Interval        =   100
      Left            =   360
      Top             =   3360
   End
   Begin VB.Frame winframe1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Multi Desk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   8
      Left            =   3120
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
      Begin VB.ListBox winlist1 
         BackColor       =   &H00D8E9EC&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2460
         Index           =   8
         ItemData        =   "main.frx":BE53
         Left            =   240
         List            =   "main.frx":BE55
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Label lblhead 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7335
   End
   Begin VB.Image Close 
      Height          =   375
      Left            =   8400
      MouseIcon       =   "main.frx":BE57
      MousePointer    =   99  'Custom
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Multi Desk"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   100
      TabIndex        =   21
      Top             =   90
      Width           =   2415
   End
   Begin VB.Menu mmnu 
      Caption         =   "MainMenu"
      Visible         =   0   'False
      Begin VB.Menu rd1 
         Caption         =   "Desk 1"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu rd2 
         Caption         =   "Desk 2"
      End
      Begin VB.Menu RD3 
         Caption         =   "Desk 3"
      End
      Begin VB.Menu rd4 
         Caption         =   "Desk 4"
      End
      Begin VB.Menu rd5 
         Caption         =   "Desk 5"
      End
      Begin VB.Menu rd6 
         Caption         =   "Desk 6"
      End
      Begin VB.Menu rd7 
         Caption         =   "Desk 7"
      End
      Begin VB.Menu rd8 
         Caption         =   "Desk 8"
      End
      Begin VB.Menu rd9 
         Caption         =   "Desk 9"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu nd 
         Caption         =   "Next Desktop"
      End
      Begin VB.Menu pd 
         Caption         =   "Previous Desktop"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu sdm 
         Caption         =   "Show Desktop Manager"
      End
      Begin VB.Menu hkc 
         Caption         =   "Hot Key && Options"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu seperator 
         Caption         =   "-"
      End
      Begin VB.Menu disable 
         Caption         =   "Disable Multi Desk"
      End
      Begin VB.Menu exitmnu 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private nic As NOTIFYICONDATA
Dim curdesk As Integer
Dim ksss As String

Private Sub about_Click()
    formAbout.Show
End Sub

''Form Events
Private Sub Form_Terminate()
    Form_Unload (1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lblhead_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
        ReleaseCapture
        SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub close_Click()
    Me.Visible = Not Me.Visible
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If False Then
        Exit Sub
    ElseIf (Me.ScaleMode = vbPixels) Then
        msg = X
        Debug.Print X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If

    Select Case msg
        Case WM_MOUSEMOVE
            txtevent.Text = "MOUSEMOVE"
        Case WM_RBUTTONDBLCLK
            txtevent.Text = "RBUTTONDBLCLK"
        Case WM_RBUTTONDOWN
            txtevent.Text = "RBUTTONDOWN"
        Case WM_RBUTTONUP
            Call SetForegroundWindow(Me.hWnd)
            Me.PopupMenu mmnu
        Case WM_MBUTTONDBLCLK
            txtevent.Text = "MBUTTONDBLCLK"
        Case WM_MBUTTONDOWN
            txtevent.Text = "MBUTTONDOWN"
        Case WM_MBUTTONUP
            txtevent.Text = "MBUTTONUP"
        Case WM_LBUTTONDBLCLK
            txtevent.Text = "LBUTTONDBLCLK"
        Case WM_LBUTTONDOWN
            Me.WindowState = vbNormal
        Case WM_LBUTTONUP
            txtevent.Text = "LBUTTONUP"
        Case Else
            txtevent.Text = "Unknown" & Str$(msg)
    End Select
End Sub

Private Sub Form_Load()
    LoadApp
End Sub
'' End of Form Events
'' Menu Events

Private Sub hkc_Click() 'show options
    configureHotKeyForm.Show
End Sub

Private Sub rd1_Click() 'desk 1
    enableDesksInMenu
    rd1.Enabled = False
    rd1.Checked = True
    movetodesktop (0)
End Sub

Private Sub rd2_Click() 'desk 2
    enableDesksInMenu
    rd2.Enabled = False
    rd2.Checked = True
    movetodesktop (1)
End Sub

Private Sub RD3_Click() 'desk 3
    enableDesksInMenu
    RD3.Enabled = False
    RD3.Checked = True
    movetodesktop (2)
End Sub

Private Sub rd4_Click() 'desk 4
    enableDesksInMenu
    rd4.Enabled = False
    rd4.Checked = True
    movetodesktop (3)
End Sub

Private Sub rd5_Click() 'desk 5
    enableDesksInMenu
    rd5.Enabled = False
    rd5.Checked = True
    movetodesktop (4)
End Sub

Private Sub rd6_Click() 'desk 6
    enableDesksInMenu
    rd6.Enabled = False
    rd6.Checked = True
    movetodesktop (5)
End Sub

Private Sub rd7_Click() 'desk 7
    enableDesksInMenu
    rd7.Enabled = False
    rd7.Checked = True
    movetodesktop (6)
End Sub

Private Sub rd8_Click() 'desk 8
    enableDesksInMenu
    rd8.Enabled = False
    rd8.Checked = True
    movetodesktop (7)
End Sub

Private Sub rd9_Click() 'desk 9
    enableDesksInMenu
    rd9.Enabled = False
    rd9.Checked = True
    movetodesktop (8)
End Sub
Private Sub nd_Click() ' next desktop
    NextDesk_Click
End Sub

Private Sub pd_Click() 'previous desktop
    PrevDesk_Click
End Sub

Private Sub sdm_Click() ' show desktop manager
    showDesktopManager
End Sub

Private Sub disable_Click() 'disable
    disableApp
End Sub

Private Sub exitmnu_Click() 'exit
    ExitApp
End Sub

'' End of Menu Events

Private Sub btndesk_Click(index As Integer)
        Select Case index
        Case 0
            rd1_Click
        Case 1
            rd2_Click
        Case 2
            RD3_Click
        Case 3
            rd4_Click
        Case 4
            rd5_Click
        Case 5
            rd6_Click
        Case 6
            rd7_Click
        Case 7
            rd8_Click
        Case 8
            rd9_Click
    End Select
End Sub

Private Sub PrevDesk_Click()
    If curdesk > 0 Then
        showDesk (-1)
    End If
End Sub

Private Sub NextDesk_Click()
    If curdesk < 8 Then
        showDesk (1)
    End If
End Sub

''Functional Code
Private Function listenToHotKey()
Dim ctrl, alt, win As Integer
        Dim cp, ap, wp As Boolean
        Dim pressed As Boolean
        cp = False
        ap = False
        wp = False
        Dim hk As String
        hk = gethk()
        'Me.Caption = hk
        ctrl = IIf(Mid(hk, 1, 1) = "1", 1, 0)
        alt = IIf(Mid(hk, 2, 1) = "1", 1, 0)
        win = IIf(Mid(hk, 3, 1) = "1", 1, 0)
    
        If ctrl = 1 Then
            cp = (GetKeyState(17) < 0)
        End If
    
        If alt = 1 Then
            ap = (GetKeyState(18) < 0)
        End If
    
        If win = 1 Then
            wp = (GetKeyState(91) < 0) Or (GetKeyState(92) < 0)
        End If
    
        press = False
        If hk = "101" Then
            press = cp And wp
        ElseIf hk = "010" Then
            press = ap
        ElseIf hk = "110" Then
            press = cp And ap
        ElseIf hk = "011" Then
            press = ap And wp
        ElseIf hk = "001" Then
            press = wp
        ElseIf hk = "100" Then
            press = cp
        ElseIf hk = "111" Then
            press = cp And ap And wp
        End If
    
        If press And (GetAsyncKeyState(39) = -32767 Or GetAsyncKeyState(40) = -32767) Then
            NextDesk_Click
        ElseIf press And (GetAsyncKeyState(37) = -32767 Or GetAsyncKeyState(38) = -32767) Then
            PrevDesk_Click
        End If

End Function

Private Function showDesktopManager()
    Me.Visible = True
    SetWindowPos Me.hWnd, IIf(True, -1, -2), 0, 0, 0, 0, &H2 Or &H1 Or &H40
    Dim u&
    u& = FillTaskListBox(winlist1(curdesk))
End Function

Private Function ExitApp()
    'show windows from all applications
    Dim hWnd&
    For i = 0 To 8
        For iss = 0 To winlist1(i).ListCount - 1
            winlist1(i).ListIndex = iss
            B = validateHidability(winlist1(i).Text, "Multi-Desk")
            If Not B Then
                hWnd& = RetHandle(winlist1(i).Text)
                u& = ShowWindow(hWnd&, SW_SHOW)
            End If
        Next
    Next
    
    showNotification 0, True 'remove tray icon
    End 'exit the app
End Function

Private Function LoadApp()

    For i = 0 To winlist1.Count - 1
    winlist1(i).Enabled = True
    Next
    Me.Visible = True
    SetWindowPos Me.hWnd, IIf(True, -1, -2), 0, 0, 0, 0, &H2 Or &H1 Or &H40 'Set window position to screen center
    Me.Visible = False
    
    'create region to crop excessive screen and drop a hole in monitor
    Dim R1, R2, R3
    R1 = CreateRectRgn(3, 25, 602, 294)
    R2 = CreateRectRgn(44, 103, 189, 220)
    R3 = CombineRgn(R1, R2, R1, 3)
    nRet = SetWindowRgn(Me.hWnd, R1, False)

    'make sure only one multidesk instance is running
    If App.PrevInstance = True Then
        MsgBox "Multi Desk is Already Launched", vbCritical, "Multi Desk"
    End
    End If
    
    App.TaskVisible = False 'hide from taskbar
    
    'respond to hotkey
    Dim hk As String
    Open App.path & "\HKC" For Input As #1
    Line Input #1, hk
    Close #1
    sethk (hk)

    'create tray icon
    curdesk = 0
    Me.Icon = LoadPicture(App.path & "\icons\enabled\" & (curdesk + 1) & ".ico")
    With nic
        .cbSize = Len(nic)
        .hWnd = Me.hWnd
        .uId = vbNull
        .uFlags = 7
        .ucallbackMessage = 512 'On Mouse Move
        .hIcon = Me.Icon
        .szTip = Me.Caption
    End With
    Shell_NotifyIcon 0, nic

End Function

Public Function validateHidability(argsrc, argfind)
    Dim h
    h = argsrc
    argsrc = Replace(argsrc, argfind, "~~~~")
    check = Not (argsrc = h)
End Function

Private Function movetodesktop(argto As Integer)
    Dim u&
    Dim hWnd&
    Dim isWhiteListed As Boolean
    u& = FillTaskListBox(winlist1(curdesk))
    
    'hide currently visible apps
    For i = 0 To winlist1(curdesk).ListCount - 1
        winlist1(curdesk).ListIndex = i
        isWhiteListed = validateHidability(winlist1(curdesk).Text, "Multi-Desk") Or validateHidability(winlist1(curdesk).Text, "Program Manager")
        If Not isWhiteListed Then
            hWnd& = RetHandle(winlist1(curdesk).Text)
            u& = ShowWindow(hWnd&, SW_HIDE)
        End If
    Next
    winframe1(curdesk).Visible = False
    
    'show apps in the desktop needed to move
    curdesk = argto
    winframe1(argto).Visible = True
    winframe1(argto).Caption = "Multi Desk " & (argto + 1)
    For i = 0 To winlist1(argto).ListCount - 1
        winlist1(argto).ListIndex = i
        isWhiteListed = validateHidability(winlist1(argto).Text, "Multi-Desk")
        If Not (isWhiteListed) Then
            hWnd& = RetHandle(winlist1(argto).Text)
            u& = ShowWindow(hWnd&, SW_SHOW)
        End If
    Next
    enableAllAndDisableCurrentDesktopInMenu (argto)
End Function

Private Function showDesk(plusOrMinus)
    Dim goIndex As Integer
    goIndex = curdesk + plusOrMinus
    movetodesktop (goIndex)
End Function

Private Function disableApp()
    disable.Caption = IIf(disable.Checked, "Disable MultiDesk", "Enable Multi Desk")
    
    'Disable
    rd1.Enabled = disable.Checked
    rd2.Enabled = disable.Checked
    RD3.Enabled = disable.Checked
    rd4.Enabled = disable.Checked
    rd5.Enabled = disable.Checked
    rd6.Enabled = disable.Checked
    rd7.Enabled = disable.Checked
    rd8.Enabled = disable.Checked
    rd9.Enabled = disable.Checked
    If (disable.Checked = True) Then
        enableAllAndDisableCurrentDesktopInMenu (curdesk)
    End If
    nd.Enabled = disable.Checked
    pd.Enabled = disable.Checked
    exitmnu.Enabled = disable.Checked
    disable.Checked = Not disable.Checked
    
    showNotification curdesk, False
End Function

Public Function enableAllAndDisableCurrentDesktopInMenu(argto)
    enableDesksInMenu
    
    Select Case argto
        Case 0
            rd1.Enabled = False
            rd1.Checked = True
        Case 1
            rd2.Enabled = False
            rd2.Checked = True
        Case 2
            RD3.Enabled = False
            RD3.Checked = True
        Case 3
            rd4.Enabled = False
            rd5.Checked = True
        Case 4
            rd5.Enabled = False
            rd5.Checked = True
        Case 5
            rd6.Enabled = False
            rd6.Checked = True
        Case 6
            rd7.Enabled = False
            rd7.Checked = True
        Case 7
            rd8.Enabled = False
            rd8.Checked = True
        Case 8
            rd9.Enabled = False
            rd9.Checked = True
    End Select
    
    showNotification argto, False

End Function


Public Function showNotification(index, remove As Boolean)
    Dim path As String
    If disable.Checked Then 'tray icon as disabed or enabled
        path = App.path & "\icons\disabled\" & (curdesk + 1) & ".ico"
    Else
        path = App.path & "\icons\enabled\" & (curdesk + 1) & ".ico"
    End If
    Me.Icon = LoadPicture(path)
    With nic
        .cbSize = Len(nic)
        .hWnd = Me.hWnd
        .uId = vbNull
        .uFlags = 7
        .ucallbackMessage = 512 'On Mouse Move
        .hIcon = Me.Icon
        .szTip = Me.Caption
    End With
    If remove Then
    Shell_NotifyIcon 0, nic
    Else
    Shell_NotifyIcon 1, nic
    End If

End Function

Private Function disableDesksInMenu()
    rd1.Enabled = False
    rd2.Enabled = False
    RD3.Enabled = False
    rd4.Enabled = False
    rd5.Enabled = False
    rd6.Enabled = False
    rd7.Enabled = False
    rd8.Enabled = False
    rd9.Enabled = False
End Function

Private Function enableDesksInMenu()
    rd1.Enabled = True
    rd1.Checked = False
    rd2.Enabled = True
    rd2.Checked = False
    RD3.Enabled = True
    RD3.Checked = False
    rd4.Enabled = True
    rd4.Checked = False
    rd5.Enabled = True
    rd5.Checked = False
    rd6.Enabled = True
    rd6.Checked = False
    rd7.Enabled = True
    rd7.Checked = False
    rd8.Enabled = True
    rd8.Checked = False
    rd9.Enabled = True
    rd9.Checked = False
End Function

Private Sub hotKeyTimer_Timer()
    If disable.Checked Then ' app disabled stop navigating
        Exit Sub
    Else
        listenToHotKey
    End If
End Sub

Private Sub winlist1_Click(index As Integer)
    winlist1(index).ToolTipText = winlist1(index).Text
End Sub
