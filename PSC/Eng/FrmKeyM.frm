VERSION 5.00
Begin VB.Form FrmKeyM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'Kein
   Caption         =   "Login screen"
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmKeyM.frx":0000
   ScaleHeight     =   3345
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox TxtPass 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   5520
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1400
      Width           =   1335
   End
   Begin VB.TextBox TxtName 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3480
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1400
      Width           =   1335
   End
   Begin VB.Image imgCancel 
      Height          =   915
      Left            =   340
      Picture         =   "FrmKeyM.frx":3C89
      Top             =   2300
      Width           =   2370
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter user and password for login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1200
      TabIndex        =   4
      Top             =   1400
      Width           =   2000
   End
   Begin VB.Image imgLogin 
      Height          =   915
      Left            =   340
      Picture         =   "FrmKeyM.frx":4739
      Top             =   120
      Width           =   2370
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5400
      TabIndex        =   3
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3350
      TabIndex        =   2
      Top             =   1050
      Width           =   1400
   End
End
Attribute VB_Name = "FrmKeyM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Module    : frmLogin
' Created:  : By Jim K on July, 2003
' Purpose   : Just a login screen for your secured application
' Use       : Use it as you want.
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CombineRgn Lib "GDI32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Sub Form_Load()
 DoTransparency Me, vbRed
End Sub
Private Sub TxtName_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
End Sub
Private Sub TxtPass_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then imgLogin_Click
End Sub
Private Sub imgLogin_Click()
 If TxtName.Text <> "" And TxtPass.Text <> "" Then
  imgLogin.Enabled = False
  imgCancel.Enabled = False
  rscon.Open "SELECT uid,urole,udata FROM dbsv_adm_user WHERE uname='" & TxtName.Text & "' AND upass='" & TxtPass.Text & "';", dbcon, adOpenDynamic, adLockOptimistic
  tmpval = ""
  If rscon.EOF = True Then
   rscon.Close
   ShowErr 0
  Else
   ChkUsrInf rscon.Fields("uid").Value, TxtName.Text, rscon.Fields("urole").Value, rscon.Fields("udata").Value
  End If
 Else
  TxtPass.SetFocus
 End If
End Sub
Private Sub imgCancel_Click()
 dbcon.Close
 Unload Me
End Sub
Private Sub ChkUsrInf(vID As Integer, vName As String, vRole As Integer, vData As Integer)
 rscon.Close
 With UsrInf
  .ID = vID
  .Name = vName
  .Role = vRole
  .Class = vData
  Select Case imgLogin.Tag
   Case 0
    If .Role = 2 Then
     .RName = "Schueler"
     rscon.Open "SELECT cname FROM dbsv_main_class WHERE cid='" & .Class & "';", dbcon, adOpenDynamic, adLockOptimistic
     tmpval = rscon.Fields("cname").Value
     rscon.Close
     Log_Entry "A", "Login DBSV: " & .Name, .ID
     Unload Me
     FrmMainS.Show
    Else
     rscon.Open "SELECT rname,acl FROM dbsv_adm_roles WHERE rid='" & .Role & "';", dbcon, adOpenDynamic, adLockOptimistic
     .RName = rscon.Fields("rname").Value
     .ACL = rscon.Fields("acl").Value
     rscon.Close
     Log_Entry "A", "Login DBSV: " & .Name, .ID
     Unload Me
     FrmMain.Show
    End If
   Case 1
    If .Role <> 1 Then
     ShowErr 0
    Else
     .RName = "Administrators"
     Log_Entry "A", "Login Admininistration: " & .Name, .ID
     Unload Me
     FrmMainA.Show
    End If
  End Select
 End With
End Sub
Private Sub ShowErr(ErrMsg As Integer)
 Log_Entry "A", "Login DBSV/Admin (Error): " & TxtName.Text, 0
 Unload Me
 FrmMsg.LoadFrm 1
End Sub
Private Function DoTransparency(bg As Form, transColor)
 Dim rgn As Long, rgn2 As Long, rgn3 As Long, rgn4 As Long
 Dim X1 As Long, Y1 As Long, i As Long, j As Long, tj As Long

 rgn = CreateRectRgn(0, 0, 0, 0)
 rgn2 = CreateRectRgn(0, 0, 0, 0)
 rgn3 = CreateRectRgn(0, 0, 0, 0)
 i = 1
 X1 = bg.Width / Screen.TwipsPerPixelX
 Y1 = bg.Height / Screen.TwipsPerPixelY
 Do While i < X1
  j = 1
  Do While j < Y1
   If GetPixel(bg.hDC, i, j) <> transColor Then
    tj = j
    Do While GetPixel(bg.hDC, i, j + 1) <> transColor
     j = j + 1
     If j = Y1 Then Exit Do
    Loop
    rgn4 = CreateRectRgn(i, tj, i + 1, j + 1)
    CombineRgn rgn3, rgn2, rgn2, 5
    CombineRgn rgn2, rgn4, rgn3, 2
    DeleteObject rgn4
   End If
   j = j + 1
  Loop
  CombineRgn rgn3, rgn, rgn, 5
  CombineRgn rgn, rgn2, rgn3, 2
  i = i + 1
 Loop
 SetWindowRgn bg.hwnd, rgn, True
 DeleteObject rgn
 DeleteObject rgn2
 DeleteObject rgn3
End Function
Function LoadFrm(Mode As Integer)
 imgLogin.Tag = Mode
 Me.Show
End Function
