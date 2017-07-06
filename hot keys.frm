VERSION 5.00
Begin VB.Form configureHotKeyForm 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   4035
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "hot keys.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "hot keys.frx":5A5A2
   ScaleHeight     =   4035
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   3480
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1700
      Left            =   600
      ScaleHeight     =   1695
      ScaleWidth      =   2175
      TabIndex        =   5
      Top             =   1120
      Width           =   2175
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1   2   3   4   5   6   7   8   9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1155
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   11925
      End
   End
   Begin VB.CheckBox ctrl 
      Caption         =   "CTRL"
      DownPicture     =   "hot keys.frx":63194
      Height          =   375
      Left            =   3360
      MouseIcon       =   "hot keys.frx":63707
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   615
   End
   Begin VB.CheckBox st 
      Caption         =   "Start Multi Desk Along with Windows"
      Height          =   375
      Left            =   3360
      MouseIcon       =   "hot keys.frx":63A11
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3120
      Value           =   2  'Grayed
      Width           =   3735
   End
   Begin VB.CheckBox win 
      Caption         =   "WIN"
      DownPicture     =   "hot keys.frx":63D1B
      Height          =   375
      Left            =   3360
      MouseIcon       =   "hot keys.frx":6428E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1755
      Width           =   615
   End
   Begin VB.CheckBox alt 
      Caption         =   "ALT"
      DownPicture     =   "hot keys.frx":64598
      Height          =   375
      Left            =   3360
      MouseIcon       =   "hot keys.frx":64B0B
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1295
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Use the key Combination choosen preceeding to the navigation keys to navigate through multiple desktops"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Image Image4 
      Enabled         =   0   'False
      Height          =   480
      Left            =   5760
      Picture         =   "hot keys.frx":64E15
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image5 
      Enabled         =   0   'False
      Height          =   480
      Left            =   6480
      Picture         =   "hot keys.frx":65257
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image6 
      Enabled         =   0   'False
      Height          =   480
      Left            =   6120
      Picture         =   "hot keys.frx":65699
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Image7 
      Enabled         =   0   'False
      Height          =   480
      Left            =   6120
      Picture         =   "hot keys.frx":65ADB
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   360
      Left            =   4800
      Picture         =   "hot keys.frx":65F1D
      Stretch         =   -1  'True
      Top             =   1290
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   6960
      MouseIcon       =   "hot keys.frx":6635F
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      ToolTipText     =   "Close"
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label lblhead 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6855
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
      TabIndex        =   10
      Top             =   90
      Width           =   2415
   End
End
Attribute VB_Name = "configureHotKeyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Dim reg As New clsRegistryAccess

Dim Cnt As Integer
Dim B As Boolean
Private Sub lblhead_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   
        ReleaseCapture
        SendMessage Me.hWnd, &HA1, 2, 0&
   
End Sub
Private Sub Form_Load()
SetWindowPos Me.hWnd, IIf(True, -1, -2), 0, 0, 0, 0, &H2 Or &H1 Or &H40
    Dim nRet As Long
    Dim R1, R2, R3
    R1 = CreateRectRgn(0, 0, 3000, 2000)
    R2 = CreateRectRgn(41, 76, 184, 187)
    R3 = CombineRgn(R1, R2, R1, 3)
    nRet = SetWindowRgn(Me.hWnd, R1, False)
Dim hk As String
hk = gethk()
ctrl.Value = IIf(Mid(hk, 1, 1) = "1", 1, 0)
alt.Value = IIf(Mid(hk, 2, 1) = "1", 1, 0)
win.Value = IIf(Mid(hk, 3, 1) = "1", 1, 0)
B = True
Dim s As String
s = ""
s = reg.ReadString("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MULTI DESK", "~~~~")

If main.validateHidability(s, "~~~") Then
st.Value = 0
Else
st.Value = 1
End If
End Sub

Private Sub Image2_Click()
Dim hk As String
hk = ""
hk = hk & IIf(ctrl.Value, 1, 0)
hk = hk & IIf(alt.Value, 1, 0)
hk = hk & IIf(win.Value, 1, 0)
Open "HKC" For Output As #1
Print #1, hk
Close #1
sethk (hk)
'MsgBox st.Value
If st.Value = 1 Then
reg.WriteString "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MULTI DESK", App.path & "\" & App.EXEName & ".exe"
Else
reg.WriteString "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MULTI DESK", "~~~~"
End If
Unload Me
End Sub

Private Sub Image3_Click()

End Sub

Private Sub Timer1_Timer()
If Label1.Left < -10500 Then
B = False
End If
If Label1.Left > 800 Then
B = True
End If
If B Then
    Label1.Left = Label1.Left - 60
    'Label2.Caption = Label1.Left
Else
    Label1.Left = Label1.Left + 60
    'Label2.Caption = Label1.Left
End If
End Sub


