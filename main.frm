VERSION 5.00
Begin VB.Form main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Multi Desk"
   ClientHeight    =   4035
   ClientLeft      =   4815
   ClientTop       =   2985
   ClientWidth     =   9015
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "main.frx":0B3A
   ScaleHeight     =   4035
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1670
      Left            =   630
      ScaleHeight     =   1665
      ScaleWidth      =   2115
      TabIndex        =   35
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
      TabIndex        =   33
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
         TabIndex        =   34
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
      TabIndex        =   31
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
         TabIndex        =   32
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
      TabIndex        =   29
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
         TabIndex        =   30
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
      TabIndex        =   27
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
         TabIndex        =   28
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
      TabIndex        =   25
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
         TabIndex        =   26
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
      TabIndex        =   23
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
         TabIndex        =   24
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
      TabIndex        =   21
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
         TabIndex        =   22
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
      TabIndex        =   19
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
         TabIndex        =   20
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
      TabIndex        =   6
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
      Begin VB.Image Image3 
         Height          =   480
         Left            =   1320
         Picture         =   "main.frx":9A75
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image2 
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
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   240
      Top             =   4200
   End
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CommandButton Command3 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   4800
         MouseIcon       =   "main.frx":BE53
         MousePointer    =   99  'Custom
         Picture         =   "main.frx":C15D
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   120
         Width           =   555
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Index           =   1
         Left            =   5520
         MouseIcon       =   "main.frx":C467
         MousePointer    =   99  'Custom
         Picture         =   "main.frx":C771
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   120
         Width           =   550
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Index           =   2
         Left            =   6240
         MouseIcon       =   "main.frx":CA7B
         MousePointer    =   99  'Custom
         Picture         =   "main.frx":CD85
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   120
         Width           =   550
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Index           =   5
         Left            =   6240
         MouseIcon       =   "main.frx":D08F
         MousePointer    =   99  'Custom
         Picture         =   "main.frx":D399
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   960
         Width           =   550
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Index           =   4
         Left            =   5520
         MouseIcon       =   "main.frx":D6A3
         MousePointer    =   99  'Custom
         Picture         =   "main.frx":D9AD
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   960
         Width           =   550
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Index           =   3
         Left            =   4800
         MouseIcon       =   "main.frx":DCB7
         MousePointer    =   99  'Custom
         Picture         =   "main.frx":DFC1
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   960
         Width           =   550
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Index           =   6
         Left            =   4800
         MouseIcon       =   "main.frx":E2CB
         MousePointer    =   99  'Custom
         Picture         =   "main.frx":E5D5
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1800
         Width           =   550
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Index           =   7
         Left            =   5520
         MouseIcon       =   "main.frx":E8DF
         MousePointer    =   99  'Custom
         Picture         =   "main.frx":EBE9
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1800
         Width           =   550
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Index           =   8
         Left            =   6240
         MaskColor       =   &H00000000&
         MouseIcon       =   "main.frx":EEF3
         MousePointer    =   99  'Custom
         Picture         =   "main.frx":F1FD
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1800
         Width           =   550
      End
      Begin VB.CommandButton Command4 
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
         Height          =   550
         Left            =   5520
         MousePointer    =   99  'Custom
         Picture         =   "main.frx":F507
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2640
         Width           =   550
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   4800
         MousePointer    =   99  'Custom
         Picture         =   "main.frx":10041
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2640
         Width           =   550
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   6240
         MousePointer    =   99  'Custom
         Picture         =   "main.frx":10483
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2640
         Width           =   550
      End
      Begin VB.TextBox txtevent 
         Height          =   495
         Left            =   1680
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   1215
      End
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   120
         Pattern         =   "*.ICO"
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8280
      Top             =   4560
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
      TabIndex        =   3
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
         ItemData        =   "main.frx":108C5
         Left            =   240
         List            =   "main.frx":108C7
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Label lblhead 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   8400
      MouseIcon       =   "main.frx":108C9
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
      TabIndex        =   36
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

Private Function movetodesktop(argto As Integer)
    Dim u&
    Dim hWnd&
    Dim isWhiteListed As Boolean
    u& = FillTaskListBox(winlist1(curdesk))
    For i = 0 To winlist1(curdesk).ListCount - 1
        winlist1(curdesk).ListIndex = i
        isWhiteListed = check(winlist1(curdesk).Text, "Multi-Desk") Or check(winlist1(curdesk).Text, "Program Manager")
        If Not isWhiteListed Then
            hWnd& = RetHandle(winlist1(curdesk).Text)
            u& = ShowWindow(hWnd&, SW_HIDE)
        End If
    Next
    winframe1(curdesk).Visible = False
    curdesk = argto
    winframe1(curdesk).Visible = True
    winframe1(curdesk).Caption = "Multi Desk " & (curdesk + 1)
    For i = 0 To winlist1(curdesk).ListCount - 1
        winlist1(curdesk).ListIndex = i
        isWhiteListed = check(winlist1(curdesk).Text, "Multi-Desk")
        If Not (isWhiteListed) Then
            hWnd& = RetHandle(winlist1(curdesk).Text)
            u& = ShowWindow(hWnd&, SW_SHOW)
        End If
    Next
    setendiss (curdesk)
End Function

Private Sub showDesk(plusOrMinus As Integer)
    Dim goIndex As Integer
    goIndex = curdesk + plusOrMinus
    movetodesktop (goIndex)
End Sub

Private Sub about_Click()
    frmSplash.Show
End Sub

Private Sub btndesk_Click(index As Integer)
    Command3_Click (index)
End Sub

Private Sub Command1_Click()
    If curdesk < 8 Then
        showDesk (1)
    End If
End Sub

Private Sub Command2_Click()
    If curdesk > 0 Then
        showDesk (-1)
    End If
End Sub

Private Sub Command3_Click(index As Integer)
    If index = 0 Then
        rd1_Click
    ElseIf index = 1 Then
        rd2_Click
    ElseIf index = 2 Then
        RD3_Click
    ElseIf index = 3 Then
        rd4_Click
    ElseIf index = 4 Then
        rd5_Click
    ElseIf index = 5 Then
        rd6_Click
    ElseIf index = 6 Then
        rd7_Click
    ElseIf index = 7 Then
        rd8_Click
    ElseIf index = 8 Then
        rd9_Click
    End If
End Sub

Private Sub disable_Click()
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
        setendiss (curdesk)
    End If
    nd.Enabled = disable.Checked
    pd.Enabled = disable.Checked
    exitmnu.Enabled = disable.Checked
    disable.Checked = Not disable.Checked
    
    'tray icon as disabed
    Dim path As String
    If disable.Checked Then
        path = App.path & "\dis\" & (curdesk + 1) & "d.ico"
    Else
        path = App.path & "\" & (curdesk + 1) & ".ico"
    End If
    Me.Icon = LoadPicture(path)
    With nic
        .cbSize = Len(nic)
        .hWnd = Me.hWnd
        .uId = vbNull
        .uFlags = 7
        .ucallbackMessage = 512 'On Mouse Move
        .hIcon = Me.Icon
        .szTip = Me.Caption + Char(0)
    End With
    Shell_NotifyIcon 1, nic
End Sub

Private Sub exitmnu_Click()
    
    'show windows from all applications
    Dim hWnd&
    For i = 0 To 8
        For iss = 0 To winlist1(i).ListCount - 1
            winlist1(i).ListIndex = iss
            B = check(winlist1(i).Text, "Multi-Desk")
            If Not B Then
                hWnd& = RetHandle(winlist1(i).Text)
                u& = ShowWindow(hWnd&, SW_SHOW)
            End If
        Next
    Next
    
    'remove tray icon
    With nic
            .cbSize = Len(nic)
            .hWnd = Me.hWnd
            .uId = vbNull
            .uFlags = 7
            .ucallbackMessage = 512 'On Mouse Move
            .hIcon = Me.Icon
            .szTip = Me.Caption + Chr(0)
    End With
    Shell_NotifyIcon 2, nic
    End
End Sub

Private Sub Form_Load()
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
    File1.path = App.path
    Line Input #1, hk
    Close #1
    sethk (hk)

    'create tray icon
    curdesk = 0
    File1.ListIndex = argto
    path = App.path & "\" & File1.FileName
    Me.Icon = LoadPicture(path)
    With nic
        .cbSize = Len(nic)
        .hWnd = Me.hWnd
        .uId = vbNull
        .uFlags = 7
        .ucallbackMessage = 512 'On Mouse Move
        .hIcon = Me.Icon
        .szTip = Me.Caption + Chr(0)
    End With
    Shell_NotifyIcon 0, nic

End Sub

Private Sub Form_Terminate()
    Form_Unload (1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
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

    'MsgBox msg
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
            'Call SetForegroundWindow(Me.hwnd)
            'Me.Show
        Case WM_LBUTTONUP
            txtevent.Text = "LBUTTONUP"
        Case Else
            txtevent.Text = "Unknown" & Str$(msg)
    End Select
End Sub

Private Sub hkc_Click()
    hotkey.Show
End Sub

Private Sub Image1_Click()
    Me.Visible = Not Me.Visible
End Sub

Private Sub Image2_Click()
    Command2_Click
End Sub

Private Sub Image3_Click()
    Command1_Click
End Sub

Private Sub lblhead_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
        ReleaseCapture
        SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub nd_Click()
    Command1_Click
End Sub

Private Sub pd_Click()
    Command2_Click
End Sub

Private Sub rd1_Click()
    enbldesks
    rd1.Enabled = False
    rd1.Checked = True
    movetodesktop (0)
End Sub

Private Sub rd2_Click()
    enbldesks
    rd2.Enabled = False
    rd2.Checked = True
    movetodesktop (1)
End Sub

Private Sub RD3_Click()
    enbldesks
    RD3.Enabled = False
    RD3.Checked = True
    movetodesktop (2)
End Sub

Private Sub rd4_Click()
    enbldesks
    rd4.Enabled = False
    rd4.Checked = True
    movetodesktop (3)
End Sub

Private Sub rd5_Click()
    enbldesks
    rd5.Enabled = False
    rd5.Checked = True
    movetodesktop (4)
End Sub

Private Sub rd6_Click()
    enbldesks
    rd6.Enabled = False
    rd6.Checked = True
    movetodesktop (5)
End Sub

Private Sub rd7_Click()
    enbldesks
    rd7.Enabled = False
    rd7.Checked = True
    movetodesktop (6)
End Sub

Private Sub rd8_Click()
    enbldesks
    rd8.Enabled = False
    rd8.Checked = True
    movetodesktop (7)
End Sub

Private Sub rd9_Click()
    enbldesks
    rd9.Enabled = False
    rd9.Checked = True
    movetodesktop (8)
End Sub

Private Sub sdm_Click()
    Me.Visible = True
    SetWindowPos Me.hWnd, IIf(True, -1, -2), 0, 0, 0, 0, &H2 Or &H1 Or &H40
    Dim u&
    u& = FillTaskListBox(winlist1(curdesk))
End Sub

Private Sub Timer2_Timer()

    If disable.Checked Then
        Exit Sub
    End If

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
        Command1_Click
    ElseIf press And (GetAsyncKeyState(37) = -32767 Or GetAsyncKeyState(38) = -32767) Then
        Command2_Click
    End If

End Sub

Private Sub winlist1_Click(index As Integer)
    winlist1(index).ToolTipText = winlist1(index).Text
End Sub

Public Function check(argsrc, argfind)
    Dim h
    h = argsrc
    argsrc = Replace(argsrc, argfind, "~~~~")
    check = Not (argsrc = h)
End Function

Private Sub disbldesks()
    rd1.Enabled = False
    rd2.Enabled = False
    RD3.Enabled = False
    rd4.Enabled = False
    rd5.Enabled = False
    rd6.Enabled = False
    rd7.Enabled = False
    rd8.Enabled = False
    rd9.Enabled = False
End Sub

Private Sub enbldesks()
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
End Sub

Public Function setendiss(argto)
    File1.ListIndex = argto
    path = App.path & "\" & File1.FileName
    enbldesks
    
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
    
    Me.Icon = LoadPicture(path)
    'curdesk = 0
    With nic
        .cbSize = Len(nic)
        .hWnd = Me.hWnd
        .uId = vbNull
        .uFlags = 7
        .ucallbackMessage = 512 'On Mouse Move
        .hIcon = Me.Icon
        .szTip = Me.Caption + Chr(0)
    End With
    Shell_NotifyIcon 1, nic
End Function
