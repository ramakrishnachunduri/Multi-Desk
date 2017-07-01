VERSION 5.00
Begin VB.Form formAbout 
   BorderStyle     =   0  'None
   ClientHeight    =   4035
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "ABOUT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ABOUT.frx":628A
   ScaleHeight     =   4035
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   3120
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   1695
      Left            =   600
      ScaleHeight     =   1695
      ScaleWidth      =   2175
      TabIndex        =   0
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
         TabIndex        =   1
         Top             =   240
         Width           =   11925
      End
   End
   Begin VB.Label lblhead 
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
      TabIndex        =   3
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"ABOUT.frx":EE7C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   2295
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   6960
      MouseIcon       =   "ABOUT.frx":EF6A
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      ToolTipText     =   "Close"
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   4
      Top             =   90
      Width           =   2415
   End
End
Attribute VB_Name = "formAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
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
B = True
End Sub

Private Sub Image2_Click()
Unload Me
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
