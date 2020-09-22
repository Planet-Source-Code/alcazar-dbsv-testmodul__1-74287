VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmUsrList 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "User Overview"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Oben ausrichten
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "addu"
            Object.ToolTipText     =   "Create User"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "editu"
            Object.ToolTipText     =   "Edit User"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "delu"
            Object.ToolTipText     =   "Delete User"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "trenn"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exit"
            Object.ToolTipText     =   "Close"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5000
      Top             =   500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUsrList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUsrList.frx":015A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUsrList.frx":0474
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUsrList.frx":078E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUsrList.frx":0AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUsrList.frx":0DC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUsrList.frx":1214
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7500
      Left            =   300
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   800
      Width           =   5300
      _ExtentX        =   9340
      _ExtentY        =   13229
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
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuusr 
      Caption         =   "User"
      Begin VB.Menu mnuaddu 
         Caption         =   "Create"
      End
      Begin VB.Menu mnueditu 
         Caption         =   "Edit"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnudelu 
         Caption         =   "Delete"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmUsrList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
 With TreeView1.Nodes
  Set nodX = .Add(, , "root", "DBSV", 1)
  rscon.Open "SELECT rid, rname FROM dbsv_adm_roles ORDER BY rid;", dbcon, adOpenDynamic, adLockOptimistic
  If rscon.EOF = False Then
   Do While rscon.EOF = False
    Set nodX = .Add("root", tvwChild, CStr(rscon.Fields("rid").Value) & "y", rscon.Fields("rname").Value, 2)
    rscon.MoveNext
   Loop
  End If
  nodX.EnsureVisible
  rscon.Close
  rscon.Open "SELECT uid, uname, urole FROM dbsv_adm_user ORDER BY uid;", dbcon, adOpenDynamic, adLockOptimistic
  If rscon.EOF = False Then
   Do While rscon.EOF = False
    Set nodX = .Add(CStr(rscon.Fields("urole").Value) & "y", tvwChild, CStr(rscon.Fields("uid").Value & "x"), rscon.Fields("uname").Value, 3)
    rscon.MoveNext
    nodX.EnsureVisible
   Loop
  End If
  rscon.Close
 End With
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
 If Node.Key = "root" Or Right$(Node.Key, 1) = "y" Then
  SetCmd False, False
 Else
  If Val(Node.Key) = UsrInf.ID Then SetCmd True, False Else SetCmd True, True
 End If
End Sub
Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
 SetCmd False, False
End Sub
Private Sub mnuclose_Click()
 Unload Me
End Sub
Private Sub mnuaddu_Click()
 FrmUsr.LoadFrm 0, 0
 SetCmd False, False
End Sub
Private Sub mnueditu_Click()
 If Val(TreeView1.SelectedItem.Key) = UsrInf.ID Then FrmUsrEdit.Show vbModal Else FrmUsr.LoadFrm 1, Val(TreeView1.SelectedItem.Key)
 SetCmd False, False
End Sub
Private Sub mnudelu_Click()
 With TreeView1
  If vbYes = MsgBox("Delete User?", vbExclamation + vbYesNo, MsgT) Then
   dbcon.Execute "DELETE FROM dbsv_adm_user WHERE uid='" & Val(.SelectedItem.Key) & "';"
   Log_Entry "D", "User deleted: " & .SelectedItem.Text, UsrInf.ID
   .Nodes.Remove (.SelectedItem.Key)
   .Refresh
  End If
 End With
 SetCmd False, False
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
  Case "addu"
   mnuaddu_Click
  Case "editu"
   mnueditu_Click
  Case "delu"
   mnudelu_Click
  Case "exit"
   mnuclose_Click
 End Select
End Sub
Private Sub SetCmd(CEdt As Boolean, CDel As Boolean)
 mnueditu.Enabled = CEdt
 Toolbar1.Buttons("editu").Enabled = CEdt
 mnudelu.Enabled = CDel
 Toolbar1.Buttons("delu").Enabled = CDel
End Sub
