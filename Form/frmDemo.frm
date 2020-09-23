VERSION 5.00
Begin VB.Form frmDemo 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo"
   ClientHeight    =   3510
   ClientLeft      =   600
   ClientTop       =   1785
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   3885
   StartUpPosition =   1  'CenterOwner
   Begin SeialKeyGenrtor.dcButton cmdVote 
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ButtonStyle     =   3
      Caption         =   "Vote Me"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SeialKeyGenrtor.dcButton cmdAbout 
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      ButtonStyle     =   3
      Caption         =   "About Serial Key Generator"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00D8E9EC&
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   480
      Width           =   3735
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thank You for Downloading Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   2760
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Noel A. Dacara (dcButton )"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   5
         Top             =   1440
         Width           =   1980
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:sabikp@gamil.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   2145
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial key Generator"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   2
         Top             =   0
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Developed by Sabith.k.p "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   2085
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================================================
'+ Serial key Generator is developed by Sabith.k.p                  +
'+ This is the demo version                                         +
'+ You may use to this code your software to protect from Softcopy  +
'+ i hope that this code help to you,pls send me your feed back     +
'+ e-mail:Sabikp@gmail.com                                          +
'====================================================================
Private Sub cmdAbout_Click()
Animate Me, 810, 4275
End Sub
Private Sub cmdVote_Click()
OpenWeb "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=66931&lngWId=1", Me.Hwnd
End Sub

Private Sub Form_Activate()
Animate Me, 810, 4275
End Sub

Private Sub Form_Load()
FrmMain.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmMain.Enabled = True
End Sub
