VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTest 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "TestModul - TestEditor"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   6810
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton CmdTest 
      Height          =   600
      Left            =   5500
      Picture         =   "FrmTest.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "run as ""Student"""
      Top             =   3600
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton CmdAusw 
      Height          =   600
      Left            =   5500
      Picture         =   "FrmTest.frx":030A
      Style           =   1  'Grafisch
      TabIndex        =   4
      ToolTipText     =   "Results"
      Top             =   2800
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton CmdClose 
      Height          =   600
      Left            =   5500
      Picture         =   "FrmTest.frx":074C
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Close"
      Top             =   4400
      Width           =   1050
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5500
      Top             =   5300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":0A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":0BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":0ED8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdAdd 
      Height          =   600
      Left            =   5500
      Picture         =   "FrmTest.frx":11F2
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Create Test"
      Top             =   400
      Width           =   1050
   End
   Begin VB.CommandButton CmdUpd 
      Height          =   600
      Left            =   5500
      Picture         =   "FrmTest.frx":1EBC
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Edit Test"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton CmdDel 
      DisabledPicture =   "FrmTest.frx":2B86
      Height          =   600
      Left            =   5500
      Picture         =   "FrmTest.frx":2F56
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Delete Test"
      Top             =   2000
      Visible         =   0   'False
      Width           =   1050
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6495
      Left            =   250
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   250
      Width           =   5000
      _ExtentX        =   8811
      _ExtentY        =   11456
      _Version        =   393217
      LabelEdit       =   1
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
   Begin VB.Menu mnudat 
      Caption         =   "File"
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnutest 
      Caption         =   "Test"
      Begin VB.Menu mnunewtest 
         Caption         =   "Create Test"
      End
      Begin VB.Menu mnuedittest 
         Caption         =   "Edit Test"
         Enabled         =   0   'False
         Begin VB.Menu mnuset 
            Caption         =   "Settings"
         End
         Begin VB.Menu mnutask 
            Caption         =   "Questions"
         End
         Begin VB.Menu mnuprint 
            Caption         =   "Print"
         End
      End
      Begin VB.Menu mnudeltest 
         Caption         =   "Delete Test"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuauswert 
         Caption         =   "Results"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnustud 
         Caption         =   "run as ""Student"""
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnutest2 
      Caption         =   "Test_Hidden"
      Visible         =   0   'False
      Begin VB.Menu mnuset2 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnutask2 
         Caption         =   "Questions"
      End
      Begin VB.Menu mnuprint2 
         Caption         =   "Print"
      End
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
 Set nodX = TreeView1.Nodes.Add(, , "root", "DBSV", 1)
 rscon.Open "SELECT DISTINCT cid,cname FROM dbsv_main_class ORDER BY cname ASC;", dbcon, adOpenDynamic, adLockOptimistic
 If rscon.EOF = False Then
  Do While rscon.EOF = False
   Set nodX = TreeView1.Nodes.Add("root", tvwChild, rscon.Fields("cid").Value & "y", rscon.Fields("cname").Value, 2)
   rscon.MoveNext
  Loop
 End If
 rscon.Close
 nodX.EnsureVisible
 If UsrInf.Role = 1 Then
  rscon.Open "SELECT tsid,tsname,tsclass,tsfach,tsactive,tsallow_online FROM dbsv_test_setting ORDER BY tsid ASC;", dbcon, adOpenDynamic, adLockOptimistic
 Else
  rscon.Open "SELECT tsid,tsname,tsclass,tsfach,tsactive,tsallow_online FROM dbsv_test_setting WHERE tsuid='" & UsrInf.ID & "' ORDER BY tsid ASC;", dbcon, adOpenDynamic, adLockOptimistic
 End If
 With TreeView1.Nodes
  If rscon.EOF = False Then
   Do While rscon.EOF = False
    Set nodX = .Add(rscon.Fields("tsclass").Value & "y", tvwChild, CStr(rscon.Fields("tsid").Value & "x"), rscon.Fields("tsname").Value, 3)
    .Item(nodX.Index).Tag = rscon.Fields("tsallow_online").Value & rscon.Fields("tsactive").Value & rscon.Fields("tsfach").Value
    rscon.MoveNext
    nodX.EnsureVisible
   Loop
  End If
 End With
 rscon.Close
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
 If Node.Key = "root" Or Right$(Node.Key, 1) = "y" Then
  SetBt False, False, False
 Else
  SetBt True, Left$(TreeView1.SelectedItem.Tag, 1), Mid$(TreeView1.SelectedItem.Tag, 2, 1)
 End If
End Sub
Private Sub mnuclose_Click()
 CmdClose_Click
End Sub
Private Sub mnunewtest_Click()
 CmdAdd_Click
End Sub
Private Sub mnuset_Click()
 FrmTestEdit1.LoadFrm 1, Val(TreeView1.SelectedItem.Key)
 SetBt False, False, False
End Sub
Private Sub mnutask_Click()
 With TreeView1.SelectedItem
  FrmTestEdit2.LoadFrm Val(.Key), .Text
 End With
 SetBt False, False, False
End Sub
Private Sub mnuprint_Click()
 MsgBox "Not available yet...", vbInformation, MsgT
 SetBt False, False, False
End Sub
Private Sub mnudeltest_Click()
 CmdDel_Click
End Sub
Private Sub mnuauswert_Click()
 CmdAusw_Click
End Sub
Private Sub mnustud_Click()
 CmdTest_Click
End Sub
Private Sub mnuset2_Click()
 mnuset_Click
End Sub
Private Sub mnutask2_Click()
 mnutask_Click
End Sub
Private Sub mnuprint2_Click()
 mnuprint_Click
End Sub
Private Sub CmdAdd_Click()
 FrmTestEdit1.LoadFrm 0, 0
 SetBt False, False, False
End Sub
Private Sub CmdUpd_Click()
 PopupMenu mnutest2, , 6700, 1200
End Sub
Private Sub CmdDel_Click()
 With TreeView1
  If vbYes = MsgBox("Do you want to delete the test " & .SelectedItem.Text & " ?" & vbCrLf & "All questions and results of the test will be deleted too.", vbExclamation + vbYesNo, MsgT) Then
   dbcon.Execute "DELETE FROM dbsv_test_result WHERE trtid='" & Val(.SelectedItem.Key) & "';"
   dbcon.Execute "DELETE FROM dbsv_test_cat WHERE tctid='" & Val(.SelectedItem.Key) & "';"
   dbcon.Execute "DELETE FROM dbsv_test_setting WHERE tsid='" & Val(.SelectedItem.Key) & "';"
   Log_Entry "D", "Deleted test: " & .SelectedItem.Text, UsrInf.ID
   .Nodes.Remove (.SelectedItem.Key)
   .Refresh
  End If
 End With
 SetBt False, False, False
End Sub
Private Sub CmdAusw_Click()
 FrmTestAuswert.LoadFrm Val(TreeView1.SelectedItem.Key), TreeView1.SelectedItem.Text
 SetBt False, False, False
End Sub
Private Sub CmdTest_Click()
 If vbYes = MsgBox("Do you want to start the selected Test?", vbExclamation + vbYesNo, MsgT) Then
  With TreeView1.SelectedItem
   FrmTestS2.LoadFrm Val(.Key), .Text, Mid$(.Tag, 3), True
  End With
 End If
 SetBt False, False, False
End Sub
Private Sub CmdClose_Click()
 Unload Me
End Sub
Private Function SetBt(Mode As Boolean, BAusw As Boolean, BStart As Boolean)
 mnuedittest.Enabled = Mode
 CmdUpd.Visible = Mode
 mnudeltest.Enabled = Mode
 CmdDel.Visible = Mode
 mnuauswert.Enabled = BAusw
 CmdAusw.Visible = BAusw
 mnuprint.Enabled = Not BAusw
 mnuprint2.Enabled = Not BAusw
 mnustud.Enabled = IIf(BAusw = True, BStart, False)
 CmdTest.Visible = IIf(BAusw = True, BStart, False)
End Function
