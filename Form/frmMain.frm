VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmMain 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serial Key Generator"
   ClientHeight    =   1695
   ClientLeft      =   690
   ClientTop       =   2835
   ClientWidth     =   6825
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   6825
   StartUpPosition =   1  'CenterOwner
   Begin SeialKeyGenrtor.dcButton CmdRegister 
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ButtonStyle     =   3
      Caption         =   "Click here to Register"
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
   Begin SeialKeyGenrtor.dcButton cmdGenrate 
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ButtonStyle     =   3
      Caption         =   "Click here to Generate Key"
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
   Begin MSComctlLib.ProgressBar PbBar 
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   609
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txtSerialkey 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtGSerailKey 
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copy the  Serial key and Paste Here"
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
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Serial Key"
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
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Key"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   840
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================================================
'+ Serial key Generator is developed by Sabithkp                    +
'+ This is the demo version                                         +
'+ You may use to this code your software to protect from Softcopy  +
'+ i hope that this code help to you,pls send me your feed back     +
'+ e-mail:Sabikp@gmail.com                                          +
'====================================================================
Option Explicit
Public SerialKey As String
Public RegVal As String
Public EncryptValueofHard  As String
Public CrackKey As String
Public SDbKey As String

Private Sub cmdGenrate_Click()
On Error GoTo Derror
    Me.Caption = "Please wait....."
            HardInfo = ""
    If rs.State = adStateOpen Then
                rs.Close
    End If
        PbBar.Value = PbBar.Value + 20
            GenerateHardDiskInfo
            SDate
        PbBar.Value = PbBar.Value + 50
            EncryptValueofHard = Encrypt(HardInfo & SCurDate)
            SerialKey = Mid(LDate, 1, 1) & Mid(EncryptValueofHard, 5, 6) & Mid(LDate, 3, 3) & Mid(EncryptValueofHard, 7, 8) & Mid(LDate, 2, 2)
            RegVal = "win" & LCase(Mid(SerialKey, 2, 9))
            SaveSetting "SKG", "SKGKey", "SCheck", RegVal
            txtGSerailKey.Text = SerialKey
                rs.Open "Select * from Skey", cn, adOpenKeyset, adLockPessimistic
                rs.Fields("key") = SerialKey 'clsp.EncryptString(SerialKey)
            rs.Update
        PbBar.Value = PbBar.Value + 30
            Randomize
    Me.Caption = "Serial Key Generator"
        cmdGenrate.Enabled = False
Exit Sub
Derror:
        If rs.State = adStateOpen Then
                rs.Close
        End If
                rs.Open "Select * from Skey", cn, adOpenKeyset, adLockPessimistic
            rs.AddNew
                rs.Fields("Key") = SerialKey 'clsp.EncryptString(SerialKey)
            rs.Update
            rs.Close
        PbBar.Value = PbBar.Value + 30
        cmdGenrate.Enabled = False
    Me.Caption = "Serial Key Generator"
End Sub

Private Sub CmdRegister_Click()
On Error GoTo RegErr
        If rs.State = adStateOpen Then
            rs.Close
        End If
                PbBar.Value = 0
            CrackKey = GetSetting("SKG", "SKGKey", "SCheck")
            rs.Open "select * from Skey", cn, adOpenKeyset, adLockPessimistic
            SDbKey = rs.Fields("Key")
        If txtSerialkey.Text = SDbKey And CrackKey = "win" & LCase(Mid(SDbKey, 2, 9)) Then
                PbBar.Value = PbBar.Value + 100
            MsgBox "Registration Completed" & vbCrLf & "      Succesfuly", vbInformation, IDentity
                txtSerialkey.Locked = True
                frmDemo.Show
        Else
            MsgBox " Invalid Serial Key " & vbCrLf & vbCrLf & "  Try Again!", vbExclamation, IDentity
            txtSerialkey.SetFocus
        End If
            rs.Close
        Exit Sub
RegErr:
    'if you erase key from Database
    MsgBox "Unhandled error has Occured" & vbCrLf & "The Current Record has been Deleted" & vbCrLf & "Generate New Key and Try Again!", vbCritical, "Unhandled error"
    End
End Sub
'
