VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmUsrEdit 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Benutzer editieren"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin TabDlg.SSTab SSTab1 
      Height          =   4000
      Left            =   300
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   300
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   7064
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Login"
      TabPicture(0)   =   "FrmUsrEdit.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "nutzid"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TxtPass"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TxtPass2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox TxtPass2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2800
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2700
         Width           =   2700
      End
      Begin VB.TextBox TxtPass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2800
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1900
         Width           =   2700
      End
      Begin MSMask.MaskEdBox nutzid 
         Height          =   360
         Left            =   2800
         TabIndex        =   0
         Top             =   1100
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   635
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "<?aaaaaaaaaaaaaaaaaaa"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Caption         =   "Passwort wiederholen:"
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
         Left            =   400
         TabIndex        =   8
         Top             =   2750
         Width           =   2100
      End
      Begin VB.Label Label2 
         Caption         =   "Passwort:"
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
         Left            =   400
         TabIndex        =   7
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "Benutzername:"
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
         Left            =   400
         TabIndex        =   6
         Top             =   1150
         Width           =   1395
      End
   End
   Begin VB.CommandButton CmdAdd 
      Height          =   660
      Left            =   1600
      Picture         =   "FrmUsrEdit.frx":001C
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   4600
      Width           =   1200
   End
   Begin VB.CommandButton CmdClose 
      Height          =   660
      Left            =   3700
      Picture         =   "FrmUsrEdit.frx":0326
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   4600
      Width           =   1200
   End
End
Attribute VB_Name = "FrmUsrEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
 nutzid.Text = UsrInf.Name
End Sub
Private Sub CmdAdd_Click()
 If nutzid.Text <> UsrInf.Name Then
  If ValID(nutzid.Text, "dbsv_adm_user", "uname") = False Then Exit Sub
  dbcon.Execute "UPDATE dbsv_adm_user SET uname='" & nutzid.Text & "' WHERE uid='" & UsrInf.ID & "';"
  UsrInf.Name = nutzid.Text
  FrmUsrList.TreeView1.Nodes(CStr(CmdAdd.Tag & "x")).Text = UsrInf.Name
  FrmUsrList.TreeView1.Refresh
  FrmMain.StatusBar1.Panels("uID").Text = "Benutzer:  " & UsrInf.Name & " ( " & UsrInf.RName & " )"
 End If
 If TxtPass.Text <> "" Then
  If IsValPw(nutzid.Text, TxtPass.Text, TxtPass2.Text) = False Then Exit Sub
  dbcon.Execute "UPDATE dbsv_adm_user SET upass='" & TxtPass2.Text & "' WHERE uid='" & UsrInf.ID & "';"
 End If
 Log_Entry "D", "Benutzer editiert: " & nutzid.Text, UsrInf.ID
 CmdClose_Click
End Sub
Private Sub CmdClose_Click()
 tmpval = ""
 Unload Me
End Sub
