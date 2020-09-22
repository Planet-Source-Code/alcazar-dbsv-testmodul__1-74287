VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRole 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Edit Role"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      Caption         =   "Permissions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1400
      Left            =   5000
      TabIndex        =   8
      Top             =   3400
      Visible         =   0   'False
      Width           =   2500
      Begin VB.OptionButton OptAcl 
         Caption         =   "Allow Access"
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
         Index           =   1
         Left            =   200
         TabIndex        =   3
         Top             =   900
         Width           =   1900
      End
      Begin VB.OptionButton OptAcl 
         Caption         =   "No Access"
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
         Index           =   0
         Left            =   200
         TabIndex        =   2
         Top             =   400
         Width           =   1500
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5500
      Left            =   300
      TabIndex        =   1
      Top             =   1400
      Width           =   4400
      _ExtentX        =   7752
      _ExtentY        =   9710
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox roleid 
      Height          =   360
      Left            =   2100
      TabIndex        =   0
      Top             =   300
      Width           =   3500
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   30
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">?<?????????????????????????????"
      PromptChar      =   " "
   End
   Begin VB.CommandButton CmdClose 
      Height          =   660
      Left            =   4100
      Picture         =   "FrmRole.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   7200
      Width           =   1200
   End
   Begin VB.CommandButton CmdAdd 
      Height          =   660
      Left            =   2400
      Picture         =   "FrmRole.frx":030A
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   7200
      Width           =   1200
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5000
      Top             =   700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRole.frx":0614
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRole.frx":076E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRole.frx":13C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRole.frx":209A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRole.frx":23B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRole.frx":26CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label26 
      Caption         =   "Objects:"
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
      Left            =   300
      TabIndex        =   7
      Top             =   1000
      Width           =   800
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Role:"
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
      Left            =   300
      TabIndex        =   6
      Top             =   350
      Width           =   1600
   End
End
Attribute VB_Name = "FrmRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TmpACL As String
Private Sub roleid_GotFocus()
 roleid.SelStart = 0
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
 Select Case Node.Key
  Case "root", "filemnu", "datamnu", "modmnu"
   Frame1.Visible = False
  Case Else
   OptAcl(Mid$(TmpACL, Val(TreeView1.SelectedItem.Key), 1)).Value = True
   Frame1.Visible = True
 End Select
End Sub
Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
 Frame1.Visible = False
End Sub
Private Sub OptAcl_Click(Index As Integer)
 Mid$(TmpACL, Val(TreeView1.SelectedItem.Key), 1) = CStr(Index)
End Sub
Private Sub CmdAdd_Click()
 tmpval = ""
 If CmdClose.Tag = 0 Then
  If ValID(roleid.Text, "dbsv_adm_roles", "rname") = False Then Exit Sub
  dbcon.Execute "INSERT INTO dbsv_adm_roles (rname,acl) VALUES('" & roleid.Text & "','" & TmpACL & "');"
  rscon.Open "SELECT last_insert_id();", dbcon, adOpenDynamic, adLockOptimistic
  tmpval = rscon.Fields("last_insert_id()").Value
  rscon.Close
  Set nodX = FrmRoleList.TreeView1.Nodes.Add("root", tvwChild, CStr(tmpval & "x"), roleid.Text, 2)
  FrmRoleList.TreeView1.Refresh
 Else
  If roleid.Text <> roleid.Tag Then
   If ValID(roleid.Text, "dbsv_adm_roles", "rname") = False Then Exit Sub
   FrmRoleList.TreeView1.Nodes(CStr(CmdAdd.Tag & "y")).Text = roleid.Text
   FrmRoleList.TreeView1.Refresh
  End If
  dbcon.Execute "UPDATE dbsv_adm_roles SET rname='" & roleid.Text & "', acl='" & TmpACL & "' WHERE rid='" & CmdAdd.Tag & "';"
 End If
 Log_Entry "D", "Edited role: " & roleid.Text, UsrInf.ID
 CmdClose_Click
End Sub
Private Sub CmdClose_Click()
 tmpval = ""
 Unload Me
End Sub
Function LoadFrm(Mode As Integer, vRID As Integer)
 CmdAdd.Tag = vRID
 CmdClose.Tag = Mode
 SetTree
 If CmdClose.Tag = 0 Then
  TmpACL = "0000"
 Else
  rscon.Open "SELECT rname,acl FROM dbsv_adm_roles WHERE rid='" & CmdAdd.Tag & "';", dbcon, adOpenDynamic, adLockOptimistic
  roleid.Text = rscon.Fields("rname").Value
  roleid.Tag = rscon.Fields("rname").Value
  TmpACL = rscon.Fields("acl").Value
  rscon.Close
 End If
 Me.Show vbModal
End Function
Private Sub SetTree()
 With TreeView1.Nodes
  Set nodX = .Add(, , "root", "DBSV", 1)
  Set nodX = .Add("root", tvwChild, "filemnu", "File", 2)
  Set nodX = .Add("root", tvwChild, "datamnu", "School Data", 2)
  Set nodX = .Add("root", tvwChild, "modmnu", "Module", 2)
  nodX.EnsureVisible
  Set nodX = .Add("filemnu", tvwChild, "1x", "LogFile", 3)
  nodX.EnsureVisible
  Set nodX = .Add("datamnu", tvwChild, "2x", "Subjects", 4)
  Set nodX = .Add("datamnu", tvwChild, "3x", "Classes", 5)
  nodX.EnsureVisible
  Set nodX = .Add("modmnu", tvwChild, "4x", "TestEditor", 6)
  nodX.EnsureVisible
 End With
End Sub
