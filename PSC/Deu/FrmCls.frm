VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCls 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Klassenliste"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5700
      Top             =   3800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCls.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCls.frx":015A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdClose 
      Height          =   660
      Left            =   5600
      Picture         =   "FrmCls.frx":0474
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Verlassen"
      Top             =   2600
      Width           =   1120
   End
   Begin VB.CommandButton CmdDel 
      Height          =   660
      Left            =   5600
      Picture         =   "FrmCls.frx":077E
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Klasse entfernen"
      Top             =   1700
      Visible         =   0   'False
      Width           =   1120
   End
   Begin VB.CommandButton CmdAdd 
      Height          =   660
      Left            =   5600
      Picture         =   "FrmCls.frx":0BC0
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Klasse erstellen"
      Top             =   800
      Width           =   1120
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5000
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   5000
      _ExtentX        =   8811
      _ExtentY        =   8811
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
      Begin VB.Menu mnunewcls 
         Caption         =   "Neue Klasse"
      End
      Begin VB.Menu mnudelcls 
         Caption         =   "Klasse entfernen"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuclose 
         Caption         =   "Schliessen"
      End
   End
End
Attribute VB_Name = "FrmCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
 With TreeView1.Nodes
  Set nodX = .Add(, , "root", "DBSV", 1)
  rscon.Open "SELECT cid,cname FROM dbsv_main_class ORDER BY cname;", dbcon, adOpenDynamic, adLockOptimistic
  If rscon.EOF = False Then
   Do While rscon.EOF = False
    Set nodX = .Add("root", tvwChild, rscon.Fields("cid").Value & "x", rscon.Fields("cname").Value, 2)
    rscon.MoveNext
   Loop
  End If
  rscon.Close
  nodX.EnsureVisible
 End With
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
 If Node.Key = "root" Then SetB False Else SetB True
End Sub
Private Sub mnunewcls_Click()
 CmdAdd_Click
End Sub
Private Sub mnudelcls_Click()
 CmdDel_Click
End Sub
Private Sub mnuclose_Click()
 CmdClose_Click
End Sub
Private Sub CmdAdd_Click()
 FrmClsDet.Show vbModal
 SetB False
End Sub
Private Sub CmdDel_Click()
 With TreeView1
  If vbYes = MsgBox("Die gewählte Klasse wirklich entfernen?", vbExclamation + vbYesNo, MsgT) Then
   dbcon.Execute "DELETE FROM dbsv_main_class WHERE cid='" & Val(.SelectedItem.Key) & "';"
   Log_Entry "D", "Klasse entfernt: " & .SelectedItem.Text, UsrInf.ID
   .Nodes.Remove (.SelectedItem.Key)
  End If
 End With
 SetB False
End Sub
Private Sub CmdClose_Click()
 Unload Me
End Sub
Private Sub SetB(Mode As Boolean)
 CmdDel.Visible = Mode
 mnudelcls.Enabled = Mode
End Sub
