VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRoleList 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Rollenübersicht"
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
            Key             =   "addr"
            Object.ToolTipText     =   "Rolle erstellen"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "editr"
            Object.ToolTipText     =   "Rolle editieren"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "delr"
            Object.ToolTipText     =   "Rolle entfernen"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "trenn"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exit"
            Object.ToolTipText     =   "Schliessen"
            ImageIndex      =   6
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRoleList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRoleList.frx":015A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRoleList.frx":0474
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRoleList.frx":078E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRoleList.frx":0AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRoleList.frx":0EFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7500
      Left            =   300
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   700
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
      Caption         =   "Datei"
      Begin VB.Menu mnuclose 
         Caption         =   "Schliessen"
      End
   End
   Begin VB.Menu mnurole 
      Caption         =   "Rollen"
      Begin VB.Menu mnuaddr 
         Caption         =   "Erstellen"
      End
      Begin VB.Menu mnueditr 
         Caption         =   "Editieren"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnudelr 
         Caption         =   "Entfernen"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmRoleList"
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
    Set nodX = .Add("root", tvwChild, CStr(rscon.Fields("rid").Value & "x"), rscon.Fields("rname").Value, 2)
    rscon.MoveNext
   Loop
  End If
  nodX.EnsureVisible
  rscon.Close
 End With
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
 Select Case Node.Key
  Case "root", "1x", "2x"
   SetCmd False, False
  Case Else
   SetCmd True, True
 End Select
End Sub
Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
 SetCmd False, False
End Sub
Private Sub mnuclose_Click()
 Unload Me
End Sub
Private Sub mnuaddr_Click()
 FrmRole.LoadFrm 0, 0
 SetCmd False, False
End Sub
Private Sub mnueditr_Click()
 FrmRole.LoadFrm 1, Val(TreeView1.SelectedItem.Key)
 SetCmd False, False
End Sub
Private Sub mnudelr_Click()
 With TreeView1
  If vbYes = MsgBox("Rolle wirklich entfernen?", vbExclamation + vbYesNo, MsgT) Then
   rscon.Open "SELECT urole FROM dbsv_adm_user WHERE urole='" & Val(.SelectedItem.Key) & "';", dbcon, adOpenDynamic, adLockOptimistic
   If rscon.EOF = False Then
    rscon.Close
    MsgBox "Rolle kann nicht entfernt werden" & vbCrLf & "Grund: Rolle ist noch Benutzern zugewiesen", vbExclamation, MsgT
   Else
    rscon.Close
    dbcon.Execute "DELETE FROM dbsv_adm_roles WHERE rid='" & Val(.SelectedItem.Key) & "';"
    Log_Entry "D", "Rolle entfernt: " & .SelectedItem.Text, UsrInf.ID
    .Nodes.Remove (.SelectedItem.Key)
    .Refresh
   End If
  End If
 End With
 SetCmd False, False
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
  Case "addr"
   mnuaddr_Click
  Case "editr"
   mnueditr_Click
  Case "delr"
   mnudelr_Click
  Case "exit"
   mnuclose_Click
 End Select
End Sub
Private Sub SetCmd(CEdt As Boolean, CDel As Boolean)
 mnueditr.Enabled = CEdt
 Toolbar1.Buttons("editr").Enabled = CEdt
 mnudelr.Enabled = CDel
 Toolbar1.Buttons("delr").Enabled = CDel
End Sub
