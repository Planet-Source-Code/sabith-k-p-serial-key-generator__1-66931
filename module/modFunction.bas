Attribute VB_Name = "modFunction"
'====================================================================
'+ Serial key Generator is developed by Sabith.k.p                  +
'+ This is the demo version                                         +
'+ You may use to this code your software to protect from Softcopy  +
'+ i hope that this code help to you,pls send me your feed back     +
'+ e-mail:Sabikp@gmail.com                                          +
'====================================================================
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal Hwnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

Public rs As New ADODB.Recordset
Public cn As New ADODB.Connection
Public HardInfo As String
Public SCurDate As String
Public LDate As String
Private obj As Object
Private obj2 As Object
Public Const IDentity = "Serial Key Generator"


Public Sub ConnectDB()
Dim dbpath As String
'On Error Resume Next
    'Get the path of the database
    dbpath = App.Path & "\Demo DataBase\Skey.mdb"
    'Open the database
    With cn
        .CommandTimeout = 5
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & dbpath & ";Persist Security Info=False"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub
Public Sub GenerateHardDiskInfo()
Dim H, hh, hhh, hhhh As Variant
On Error GoTo driveErr:
Set obj = GetObject("winmgmts:").InstancesOf("Win32_DiskDrive")
            For Each obj2 In obj
             H = obj2.Size
            hh = H / 1024
            hhh = hh / 1024
            hhhh = hhh / 1024
           HardInfo = HardInfo & obj2.Caption
Next
  Exit Sub
driveErr:
  MsgBox HardInfo & "Removable Drive", vbInformation, IDentity
End Sub

Public Function Encrypt(StringToEncrypt As String, Optional AlphaEncoding As Boolean = False) As String
    On Error GoTo ErrorHandler
    Dim Char As String
    Dim i As Integer
    
    Encrypt = ""
    
    For i = 1 To Len(StringToEncrypt)
        Char = Asc(Mid(StringToEncrypt, i, 1))
        Encrypt = Encrypt & Len(Char) & Char
    Next i
    
    If AlphaEncoding Then
        
        StringToEncrypt = Encrypt
        Encrypt = ""


        For i = 1 To Len(StringToEncrypt)
            Encrypt = Encrypt & Chr(Mid(StringToEncrypt, i, 1) + 19129)
        Next i
    End If
    Exit Function
ErrorHandler:
    Encrypt = "Error encrypting string"
End Function

Public Function Decrypt(StringToDecrypt As String, Optional AlphaDecoding As Boolean = False) As String
    On Error GoTo ErrorHandler
    Dim CharCode As String
    Dim CharPos As Integer
    Dim Char As String
    Dim i As Integer
    

    If AlphaDecoding Then
        
        Decrypt = StringToDecrypt
        StringToDecrypt = ""


        For i = 1 To Len(Decrypt)
            StringToDecrypt = StringToDecrypt & (Asc(Mid(Decrypt, i, 1)) - 19129)
        Next i
    End If
    
   Decrypt = ""
    
    Do
        
        CharPos = Left(StringToDecrypt, 1)
        StringToDecrypt = Mid(StringToDecrypt, 2)
        CharCode = Left(StringToDecrypt, CharPos)
       StringToDecrypt = Mid(StringToDecrypt, Len(CharCode) + 1)
        Decrypt = Decrypt & Chr(CharCode)
    Loop Until StringToDecrypt = ""
    Exit Function
ErrorHandler:
    Decrypt = "Error decrypting string"
      MsgBox "Value Not Found ", vbExclamation, IDentity
End Function
Public Function SDate()
Dim GDate As String
LDate = UCase(Format(Date, "Long Date"))
GDate = Mid(LDate, 1, 3)
SCurDate = GDate
End Function
Public Function Animate(frm As Form, hMin As Long, hMax As Long)
Dim i As Integer
For i = hMin To hMax
frm.Height = i
Next i
End Function
Sub Main()
ConnectDB
FrmMain.Show
End Sub
Public Sub OpenWeb(webAdd As String, sHWND As Long)
      Call ShellExecute(sHWND, vbNullString, webAdd, "", vbNullString, 1)
End Sub
