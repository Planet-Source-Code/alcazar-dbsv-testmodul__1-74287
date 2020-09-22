VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTestS 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "TestModul - Tests"
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
   Begin VB.CommandButton CmdAusw 
      Height          =   600
      Left            =   5500
      Picture         =   "FrmTestS.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Show Results"
      Top             =   1600
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton CmdStart 
      Height          =   600
      Left            =   5500
      Picture         =   "FrmTestS.frx":0442
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Start Test"
      Top             =   800
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton CmdClose 
      Height          =   600
      Left            =   5500
      Picture         =   "FrmTestS.frx":074C
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Close"
      Top             =   2400
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
            Picture         =   "FrmTestS.frx":0A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTestS.frx":0BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTestS.frx":0ED8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6495
      Left            =   250
      TabIndex        =   3
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
      Begin VB.Menu mnustart 
         Caption         =   "Start Test"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuauswert 
         Caption         =   "Show Results"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmTestS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstmp As String
Private Sub Form_Load()
 With TreeView1.Nodes
  Set nodX = .Add(, , "root", "DBSV", 1)
  rscon.Open "SELECT f.fid, f.name FROM dbsv_main_fach AS f JOIN dbsv_test_setting AS ts ON f.fid=ts.tsfach WHERE ts.tsclass='" & UsrInf.Class & "' AND ts.tsallow_online='1';", dbcon, adOpenDynamic, adLockOptimistic
  If rscon.EOF = False Then
   Do While rscon.EOF = False
    Set nodX = .Add("root", tvwChild, rscon.Fields("fid").Value & "y", rscon.Fields("name").Value, 2)
    rscon.MoveNext
   Loop
  End If
  rscon.Close
  nodX.EnsureVisible
  rscon.Open "SELECT tsid,tsname,tsfach FROM dbsv_test_setting WHERE tsclass='" & UsrInf.Class & "' AND tsactive='1' AND tsallow_online='1' ORDER BY tsid ASC;", dbcon, adOpenDynamic, adLockOptimistic
  If rscon.EOF = False Then
   Do While rscon.EOF = False
    Set nodX = .Add(rscon.Fields("tsfach").Value & "y", tvwChild, CStr(rscon.Fields("tsid").Value & "x"), rscon.Fields("tsname").Value, 3)
    nodX.EnsureVisible
    rscon.MoveNext
   Loop
  End If
  rscon.Close
 End With
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
 If Node.Key = "root" Or Right$(Node.Key, 1) = "y" Then
  SetBt False
 Else
  SetBt True
 End If
End Sub
Private Sub mnuclose_Click()
 CmdClose_Click
End Sub
Private Sub mnustart_Click()
 CmdStart_Click
End Sub
Private Sub mnuauswert_Click()
 CmdAusw_Click
End Sub
Private Sub CmdStart_Click()
 rstmp = ""
 tmpval = ""
 rscon.Open "SELECT tsmultilimit,tsdelay FROM dbsv_test_setting WHERE tsclass='" & UsrInf.Class & "' AND tsfach='" & Val(TreeView1.SelectedItem.Parent.Key) & "';", dbcon, adOpenDynamic, adLockOptimistic
 rstmp = rscon.Fields("tsmultilimit").Value
 tmpval = rscon.Fields("tsdelay").Value
 rscon.Close
 rscon.Open "SELECT COUNT(*) AS anzahl, trdatum FROM dbsv_test_result WHERE truid='" & UsrInf.ID & "' AND trtid='" & Val(TreeView1.SelectedItem.Key) & "';", dbcon, adOpenDynamic, adLockOptimistic
 If Val(rstmp) > 1 Then
  If rscon.Fields("anzahl").Value >= Val(rstmp) Then
   ShowMsg "Your maximum tries for this test have been reached"
   Exit Sub
  End If
  If Val(tmpval) > 0 And Val(ConvertToTimeStamp(Now())) < (ConvertToTimeStamp(rscon.Fields("trdatum").Value) + Val(tmpval) * 60) Then
   ShowMsg "This test has a delay of " & tmpval & " minutes between tries" & vbCrLf & "and this time hasnt run out yet"
   Exit Sub
  End If
  rscon.Close
  With TreeView1.SelectedItem
   FrmTestS2.LoadFrm Val(.Key), .Text, Val(.Parent.Key), False
  End With
 Else
  If rscon.Fields("anzahl").Value >= 1 Then
   ShowMsg "This Test can be only taken once per user" & vbCrLf & "You have already taken it"
   Exit Sub
  End If
  rscon.Close
  With TreeView1.SelectedItem
   FrmTestS2.LoadFrm Val(.Key), .Text, Val(.Parent.Key), False
  End With
 End If
 SetBt False
End Sub
Private Sub CmdAusw_Click()
 FrmTestAuswert.LoadFrm Val(TreeView1.SelectedItem.Key), TreeView1.SelectedItem.Text
 SetBt False
End Sub
Private Sub CmdClose_Click()
 Unload Me
End Sub
Private Function SetBt(Mode As Boolean)
 mnustart.Enabled = Mode
 CmdStart.Visible = Mode
 mnuauswert.Enabled = Mode
 CmdAusw.Visible = Mode
End Function
Private Function ShowMsg(MsgTxt As String)
 rscon.Close
 rstmp = ""
 tmpval = ""
 MsgBox MsgTxt & vbCrLf & "Test is not started", vbExclamation, MsgT
 SetBt False
End Function
Private Function ConvertToTimeStamp(ToCDate As String) As String
 ConvertToTimeStamp = DateDiff("s", "01.01.1970 00:00:00", ToCDate, 2, 1)
End Function
