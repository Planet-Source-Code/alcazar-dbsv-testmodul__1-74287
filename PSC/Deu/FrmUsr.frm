VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmUsr 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Edit User"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin TabDlg.SSTab SSTab1 
      Height          =   4000
      Left            =   300
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   300
      Width           =   6500
      _ExtentX        =   11456
      _ExtentY        =   7064
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabPicture(0)   =   "FrmUsr.frx":0000
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
      TabCaption(1)   =   "Daten"
      TabPicture(1)   =   "FrmUsr.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblStud"
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(2)=   "Combo2"
      Tab(1).Control(3)=   "Combo1"
      Tab(1).ControlCount=   4
      Begin VB.ComboBox Combo1 
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
         Left            =   -72500
         Style           =   2  'Dropdown-Liste
         TabIndex        =   3
         Top             =   1100
         Width           =   2700
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   -72500
         Style           =   2  'Dropdown-Liste
         TabIndex        =   4
         Top             =   2000
         Visible         =   0   'False
         Width           =   2700
      End
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
         Left            =   3000
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
         Left            =   3000
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1900
         Width           =   2700
      End
      Begin MSMask.MaskEdBox nutzid 
         Height          =   360
         Left            =   3000
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
      Begin VB.Label Label1 
         Caption         =   "Rolle:"
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
         Left            =   -74400
         TabIndex        =   12
         Top             =   1150
         Width           =   600
      End
      Begin VB.Label LblStud 
         Caption         =   "Klasse:"
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
         Left            =   -74400
         TabIndex        =   11
         Top             =   2050
         Visible         =   0   'False
         Width           =   795
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
         Left            =   600
         TabIndex        =   10
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
         Left            =   600
         TabIndex        =   9
         Top             =   1950
         Width           =   1000
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
         Left            =   600
         TabIndex        =   8
         Top             =   1150
         Width           =   1395
      End
   End
   Begin VB.CommandButton CmdClose 
      Height          =   660
      Left            =   4000
      Picture         =   "FrmUsr.frx":0038
      Style           =   1  'Grafisch
      TabIndex        =   6
      Top             =   4600
      Width           =   1200
   End
   Begin VB.CommandButton CmdAdd 
      Height          =   660
      Left            =   2000
      Picture         =   "FrmUsr.frx":0342
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   4600
      Width           =   1200
   End
End
Attribute VB_Name = "FrmUsr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TmpRole As Integer, zc As Integer
Private Sub nutzid_GotFocus()
 nutzid.SelStart = 0
End Sub
Private Sub Combo1_Click()
 If Combo1.ItemData(Combo1.ListIndex) = 2 Then SetCmb True Else SetCmb False
End Sub
Private Sub CmdAdd_Click()
 tmpval = ""
 TmpRole = Combo1.ItemData(Combo1.ListIndex)
 If CmdClose.Tag = 0 Then
  If ValID(nutzid.Text, "dbsv_adm_user", "uname") = False Then Exit Sub
  If IsValPw(nutzid.Text, TxtPass.Text, TxtPass2.Text) = False Then Exit Sub
  If TmpRole = 2 Then
   If Combo2.ListIndex = -1 Then
    MsgBox "Wählen Sie die Klasse des Schülerbenutzers aus", vbExclamation, MsgT
    Exit Sub
   End If
   zc = Combo2.ItemData(Combo2.ListIndex)
  Else
   zc = 0
  End If
  dbcon.Execute "INSERT INTO dbsv_adm_user (uname,upass,urole,udata) VALUES('" & nutzid.Text & "','" & TxtPass2.Text & "','" & TmpRole & "','" & zc & "');"
  rscon.Open "SELECT last_insert_id();", dbcon, adOpenDynamic, adLockOptimistic
  tmpval = rscon.Fields("last_insert_id()").Value
  rscon.Close
  Set nodX = FrmUsrList.TreeView1.Nodes.Add(CStr(TmpRole & "y"), tvwChild, CStr(tmpval & "x"), nutzid.Text, 3)
  FrmUsrList.TreeView1.Refresh
 Else
  If nutzid.Text <> nutzid.Tag Then
   If ValID(nutzid.Text, "dbsv_adm_user", "uname") = False Then Exit Sub
   FrmUsrList.TreeView1.Nodes(CStr(CmdAdd.Tag & "x")).Text = nutzid.Text
   FrmUsrList.TreeView1.Refresh
  End If
  If TxtPass.Text <> "" Then
   If IsValPw(nutzid.Text, TxtPass.Text, TxtPass2.Text) = False Then Exit Sub
  End If
  If TxtPass.Text <> "" Then
   dbcon.Execute "UPDATE dbsv_adm_user SET uname='" & nutzid.Text & "', upass='" & TxtPass2.Text & "', urole='" & TmpRole & "' WHERE uid='" & CmdAdd.Tag & "';"
  Else
   dbcon.Execute "UPDATE dbsv_adm_user SET uname='" & nutzid.Text & "', urole='" & TmpRole & "' WHERE uid='" & CmdAdd.Tag & "';"
  End If
 End If
 Log_Entry "D", "Benutzer editiert: " & nutzid.Text, UsrInf.ID
 CmdClose_Click
End Sub
Private Sub CmdClose_Click()
 tmpval = ""
 Unload Me
End Sub
Function LoadFrm(Mode As Integer, vUsr As Integer)
 CmdAdd.Tag = vUsr
 CmdClose.Tag = Mode
 rscon.Open "SELECT rid,rname FROM dbsv_adm_roles ORDER BY rid", dbcon, adOpenDynamic, adLockOptimistic
 If rscon.EOF = False Then
  Do While rscon.EOF = False
   If rscon.Fields("rid").Value = 2 Then
    If CmdClose.Tag = 0 Then
     Combo1.AddItem rscon.Fields("rname").Value
     Combo1.ItemData(Combo1.NewIndex) = rscon.Fields("rid").Value
    End If
   Else
    Combo1.AddItem rscon.Fields("rname").Value
    Combo1.ItemData(Combo1.NewIndex) = rscon.Fields("rid").Value
   End If
   rscon.MoveNext
  Loop
 End If
 rscon.Close
 rscon.Open "SELECT cid, cname from dbsv_main_class ORDER BY cname;", dbcon, adOpenDynamic, adLockOptimistic
 If rscon.EOF = False Then
  Do While rscon.EOF = False
   Combo2.AddItem rscon.Fields("cname").Value
   Combo2.ItemData(Combo2.NewIndex) = rscon.Fields("cid").Value
   rscon.MoveNext
  Loop
 End If
 rscon.Close
 If CmdClose.Tag = 0 Then
  Combo1.ListIndex = 1
  Combo2.ListIndex = 0
 Else
  rscon.Open "SELECT a.uname,a.urole,b.rname FROM dbsv_adm_user AS a JOIN dbsv_adm_roles AS b ON a.urole=b.rid WHERE a.uid='" & CmdAdd.Tag & "';", dbcon, adOpenDynamic, adLockOptimistic
  nutzid.Text = rscon.Fields("uname").Value
  nutzid.Tag = nutzid.Text
  If rscon.Fields("urole").Value = "2" Then
   With Combo1
    .Clear
    .AddItem "Schueler"
    .ItemData(.NewIndex) = 2
    .ListIndex = 0
   End With
   SSTab1.TabEnabled(1) = False
  Else
   Combo1.Text = rscon.Fields("rname").Value
  End If
  rscon.Close
 End If
 Me.Show vbModal
End Function
Private Sub SetCmb(MStud As Boolean)
 LblStud.Visible = MStud
 Combo2.Visible = MStud
End Sub
