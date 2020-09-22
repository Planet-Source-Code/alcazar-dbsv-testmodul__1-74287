VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmTestEdit1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "TestEditor - Einstellungen"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton CmdSave 
      Height          =   660
      Left            =   2400
      Picture         =   "FrmTestEdit1.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   11
      ToolTipText     =   "Einstellungen speichern"
      Top             =   6000
      Width           =   1320
   End
   Begin VB.CommandButton CmdClose 
      Height          =   660
      Left            =   4320
      Picture         =   "FrmTestEdit1.frx":030A
      Style           =   1  'Grafisch
      TabIndex        =   12
      ToolTipText     =   "Verlassen"
      Top             =   6000
      Width           =   1320
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5400
      Left            =   300
      TabIndex        =   13
      Top             =   300
      Width           =   7300
      _ExtentX        =   12885
      _ExtentY        =   9525
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
      TabCaption(0)   =   "Allgemein"
      TabPicture(0)   =   "FrmTestEdit1.frx":0614
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label19"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TxtName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmbClass"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CmbFach"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TxtIntro"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Online-Test"
      TabPicture(1)   =   "FrmTestEdit1.frx":0630
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ChkTstOnline"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3900
         Left            =   -74400
         TabIndex        =   18
         Top             =   1100
         Visible         =   0   'False
         Width           =   4300
         Begin VB.CheckBox ChkShowQ 
            Alignment       =   1  'Rechts ausgerichtet
            Caption         =   "Aufgaben sind auswählbar"
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
            TabIndex        =   9
            Top             =   2700
            Width           =   2900
         End
         Begin VB.CheckBox ChkExp 
            Alignment       =   1  'Rechts ausgerichtet
            Caption         =   "Bei Zeitablauf ""nicht bestanden"""
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
            Top             =   950
            Width           =   3400
         End
         Begin MSMask.MaskEdBox TxtScore 
            Height          =   375
            Left            =   2200
            TabIndex        =   10
            Top             =   3300
            Width           =   550
            _ExtentX        =   953
            _ExtentY        =   661
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtDelay 
            Height          =   360
            Left            =   2500
            TabIndex        =   8
            Top             =   2100
            Width           =   650
            _ExtentX        =   1138
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "9999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtMulti 
            Height          =   360
            Left            =   2000
            TabIndex        =   7
            Top             =   1500
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "99"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtTime 
            Height          =   360
            Left            =   1700
            TabIndex        =   5
            Top             =   300
            Width           =   550
            _ExtentX        =   979
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "999"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            Caption         =   "%"
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
            Left            =   2850
            TabIndex        =   26
            Top             =   3350
            Width           =   300
         End
         Begin VB.Label Label8 
            Caption         =   "Pause zwischen Tests:"
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
            TabIndex        =   25
            Top             =   2150
            Width           =   2100
         End
         Begin VB.Label Label7 
            Caption         =   "Teilnahme-Limit:"
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
            TabIndex        =   24
            Top             =   1550
            Width           =   1600
         End
         Begin VB.Label Label6 
            Caption         =   "Test-Zeitlimit:"
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
            TabIndex        =   23
            Top             =   350
            Width           =   1300
         End
         Begin VB.Label Label3 
            Caption         =   "Test bestanden ab:"
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
            TabIndex        =   22
            Top             =   3350
            Width           =   1800
         End
         Begin VB.Label Label15 
            Caption         =   "Minuten"
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
            Left            =   2400
            TabIndex        =   21
            Top             =   350
            Width           =   700
         End
         Begin VB.Label Label16 
            Caption         =   "x"
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
            Left            =   2550
            TabIndex        =   20
            Top             =   1550
            Width           =   195
         End
         Begin VB.Label Label18 
            Caption         =   "Minuten"
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
            Left            =   3300
            TabIndex        =   19
            Top             =   2150
            Width           =   705
         End
      End
      Begin VB.CheckBox ChkTstOnline 
         Caption         =   "Online-Teilnahme erlauben"
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
         Left            =   -74280
         TabIndex        =   4
         Top             =   800
         Width           =   2800
      End
      Begin VB.TextBox TxtIntro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   2100
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2800
         Width           =   3600
      End
      Begin VB.ComboBox CmbFach 
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
         Left            =   2100
         Style           =   2  'Dropdown-Liste
         TabIndex        =   2
         Top             =   2100
         Width           =   2500
      End
      Begin VB.ComboBox CmbClass 
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
         Left            =   2100
         Style           =   2  'Dropdown-Liste
         TabIndex        =   1
         Top             =   1500
         Width           =   2500
      End
      Begin VB.TextBox TxtName 
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
         Left            =   2100
         MaxLength       =   30
         TabIndex        =   0
         Top             =   800
         Width           =   3600
      End
      Begin VB.Label Label19 
         Caption         =   "Einleitung"
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
         TabIndex        =   17
         Top             =   2800
         Width           =   1000
      End
      Begin VB.Label Label2 
         Caption         =   "Fach"
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
         TabIndex        =   16
         Top             =   2150
         Width           =   500
      End
      Begin VB.Label Label4 
         Caption         =   "Klasse"
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
         TabIndex        =   15
         Top             =   1550
         Width           =   700
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
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
         TabIndex        =   14
         Top             =   850
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmTestEdit1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tid As Integer
Private Sub TxtTime_GotFocus()
 TxtTime.SelStart = 0
End Sub
Private Sub TxtMulti_GotFocus()
 TxtMulti.SelStart = 0
End Sub
Private Sub TxtDelay_GotFocus()
 TxtDelay.SelStart = 0
End Sub
Private Sub TxtScore_GotFocus()
 TxtScore.SelStart = 0
End Sub
Private Sub TxtTime_Change()
 If TxtTime.Text = "" Or Val(TxtTime.Text) > 240 Then
  MsgBox "Das Zeitlimit muss zwischen 0 und 240 Minuten liegen", vbExclamation, MsgT
  TxtTime.Text = 60
 End If
End Sub
Private Sub TxtMulti_Change()
 If TxtMulti.Text = "" Or Val(TxtMulti.Text) = 0 Then
  MsgBox "Geben Sie für Mehrfach-Teilnahme einen Wert größer 0 an", vbExclamation, MsgT
  TxtMulti.Text = 1
 End If
End Sub
Private Sub TxtDelay_Change()
 If TxtDelay.Text = "" Or Val(TxtDelay.Text) > 1440 Then
  MsgBox "Das Zeitlimit muss zwischen 0 und 1440 Minuten liegen", vbExclamation, MsgT
  TxtDelay.Text = 30
 End If
End Sub
Private Sub TxtScore_Change()
 If TxtScore.Text = "" Or Val(TxtScore.Text) = 0 Or Val(TxtScore.Text) > 100 Then
  MsgBox "Die Prozentwertung muss zwischen 1 und 100 liegen", vbExclamation, MsgT
  TxtScore.Text = 50
 End If
End Sub
Private Sub ChkTstOnline_Click()
 Frame2.Visible = IIf(ChkTstOnline.Value = 1, 1, 0)
End Sub
Private Sub CmdSave_Click()
 If CmdSave.Tag = 0 Then
  If ValID(TxtName.Text, "dbsv_test_setting", "tsname") = False Then Exit Sub
 Else
  If TxtName.Text <> TxtName.Tag Then
   If ValID(TxtName.Text, "dbsv_test_setting", "tsname") = False Then Exit Sub
  End If
 End If
 If CmbClass.ListIndex = -1 Or CmbFach.ListIndex = -1 Then
  MsgBox "Wählen Sie die Klasse und das Fach aus für die der Test erstellt wird", vbExclamation, MsgT
  Exit Sub
 End If
 If TxtIntro.Text = "" Then
  MsgBox "Geben Sie eine einleitende Beschreibung an", vbExclamation, MsgT
  Exit Sub
 End If
 If CmdSave.Tag = 0 Then
  tmpval = ""
  dbcon.Execute "INSERT INTO dbsv_test_setting (tsuid, tsname, tsintro, tsclass, tsfach, tsactive, tsallow_online, tstimelimit, tstime_exp, tsmultilimit, tsdelay, tsshowq, tsscore_pass) VALUES ('" & UsrInf.ID & "','" & TxtName.Text & "','" & TxtIntro.Text & "','" & CmbClass.ItemData(CmbClass.ListIndex) & "','" & CmbFach.ItemData(CmbFach.ListIndex) & "','0','" & ChkTstOnline.Value & "','" & TxtTime.Text & "','" & ChkExp.Value & "','" & TxtMulti.Text & "','" & TxtDelay.Text & "','" & ChkShowQ.Value & "','" & TxtScore.Text & "');"
  rscon.Open "SELECT last_insert_id();", dbcon, adOpenDynamic, adLockOptimistic
  tmpval = rscon.Fields("last_insert_id()").Value
  rscon.Close
  With FrmTest.TreeView1
   Set nodX = .Nodes.Add(CmbClass.ItemData(CmbClass.ListIndex) & "y", tvwChild, tmpval & "x", TxtName.Text, 3)
   .Nodes.Item(tmpval & "x").Tag = ChkTstOnline.Value
   nodX.EnsureVisible
   .Refresh
  End With
 Else
  dbcon.Execute "UPDATE dbsv_test_setting SET tsname='" & TxtName.Text & "', tsintro='" & TxtIntro.Text & "', tsallow_online='" & ChkTstOnline.Value & "', tstimelimit='" & TxtTime.Text & "', tstime_exp='" & ChkExp.Value & "', tsmultilimit='" & TxtMulti.Text & "', tsdelay='" & TxtDelay.Text & "', tsshowq='" & ChkShowQ.Value & "', tsscore_pass='" & TxtScore.Text & "' WHERE tsid='" & tid & "';"
  If TxtName.Text <> TxtName.Tag Then
   FrmTest.TreeView1.Nodes(CStr(tid & "x")).Text = TxtName.Text
   FrmTest.TreeView1.Refresh
  End If
 End If
 Log_Entry "D", "Test erstellt/editiert: " & TxtName.Text, UsrInf.ID
 CmdClose_Click
End Sub
Private Sub CmdClose_Click()
 tmpval = ""
 Unload Me
End Sub
Function LoadFrm(Mode As Integer, vTest As Integer)
 CmdSave.Tag = Mode
 tid = vTest
 If CmdSave.Tag = 1 Then
  rscon.Open "SELECT t.*, c.cid, c.cname, f.fid, f.name AS fname FROM dbsv_test_setting AS t JOIN dbsv_main_class AS c ON t.tsclass=c.cid JOIN dbsv_main_fach AS f ON t.tsfach=f.fid WHERE t.tsid='" & tid & "' AND c.cid=t.tsclass AND f.fid=t.tsfach;", dbcon, adOpenDynamic, adLockOptimistic
  TxtName.Text = rscon.Fields("tsname").Value
  TxtName.Tag = rscon.Fields("tsname").Value
  CmbClass.AddItem rscon.Fields("cname").Value
  CmbClass.ItemData(CmbClass.NewIndex) = rscon.Fields("cid").Value
  CmbClass.ListIndex = 0
  CmbFach.AddItem rscon.Fields("fname").Value
  CmbFach.ItemData(CmbFach.NewIndex) = rscon.Fields("fid").Value
  CmbFach.ListIndex = 0
  TxtIntro.Text = rscon.Fields("tsintro").Value
  ChkTstOnline.Value = rscon.Fields("tsallow_online").Value
  TxtTime.Text = rscon.Fields("tstimelimit").Value
  ChkExp.Value = rscon.Fields("tstime_exp").Value
  TxtMulti.Text = rscon.Fields("tsmultilimit").Value
  TxtDelay.Text = rscon.Fields("tsdelay").Value
  ChkShowQ.Value = rscon.Fields("tsshowq").Value
  TxtScore.Text = rscon.Fields("tsscore_pass").Value
  rscon.Close
  Label4.Enabled = False
  Label2.Enabled = False
  CmbClass.Enabled = False
  CmbFach.Enabled = False
 Else
  rscon.Open "SELECT cid,cname FROM dbsv_main_class ORDER BY cname ASC;", dbcon, adOpenDynamic, adLockOptimistic
  If rscon.EOF = False Then
   Do While rscon.EOF = False
    CmbClass.AddItem rscon.Fields("cname").Value
    CmbClass.ItemData(CmbClass.NewIndex) = rscon.Fields("cid").Value
    rscon.MoveNext
   Loop
  End If
  rscon.Close
  rscon.Open "SELECT fid,name FROM dbsv_main_fach ORDER BY name ASC;", dbcon, adOpenDynamic, adLockOptimistic
  If rscon.EOF = False Then
   Do While rscon.EOF = False
    CmbFach.AddItem rscon.Fields("name").Value
    CmbFach.ItemData(CmbFach.NewIndex) = rscon.Fields("fid").Value
    rscon.MoveNext
   Loop
  End If
  rscon.Close
  TxtTime.Text = 60
  ChkExp.Value = 0
  TxtMulti.Text = 1
  TxtDelay.Text = 30
  ChkShowQ.Value = 0
  TxtScore.Text = 50
 End If
 Me.Show vbModal
End Function
