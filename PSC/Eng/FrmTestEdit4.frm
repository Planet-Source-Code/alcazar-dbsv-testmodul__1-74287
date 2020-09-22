VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmTestEdit4 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "TestEditor - Edit Questions"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox TxtA1E 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1500
      MaxLength       =   200
      TabIndex        =   2
      Top             =   3700
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.CheckBox CheckCA 
      Height          =   300
      Index           =   2
      Left            =   9300
      TabIndex        =   9
      Top             =   6050
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CheckBox CheckCA 
      Height          =   300
      Index           =   3
      Left            =   9300
      TabIndex        =   10
      Top             =   7200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CheckBox CheckCA 
      Height          =   300
      Index           =   4
      Left            =   9300
      TabIndex        =   11
      Top             =   8350
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CheckBox CheckCA 
      Height          =   255
      Index           =   1
      Left            =   9300
      TabIndex        =   8
      Top             =   4900
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'Kein
      Height          =   3400
      Left            =   350
      TabIndex        =   26
      Top             =   5650
      Width           =   8750
      Begin VB.TextBox TxtA5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   950
         Left            =   1150
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   2350
         Width           =   7500
      End
      Begin VB.TextBox TxtA4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   950
         Left            =   1150
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   7500
      End
      Begin VB.TextBox TxtA3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   950
         Left            =   1150
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   50
         Width           =   7500
      End
      Begin VB.Label Label7 
         Caption         =   "Answer 5:"
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
         Left            =   50
         TabIndex        =   30
         Top             =   2700
         Width           =   900
      End
      Begin VB.Label Label6 
         Caption         =   "Answer 4:"
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
         Left            =   50
         TabIndex        =   29
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label Label5 
         Caption         =   "Answer 3:"
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
         Left            =   50
         TabIndex        =   28
         Top             =   400
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Kein
      Height          =   1100
      Left            =   350
      TabIndex        =   25
      Top             =   4500
      Width           =   8750
      Begin VB.TextBox TxtA2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   950
         Left            =   1150
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   50
         Width           =   7500
      End
      Begin VB.Label Label1 
         Caption         =   "Answer 2:"
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
         Left            =   50
         TabIndex        =   27
         Top             =   400
         Width           =   900
      End
   End
   Begin VB.CheckBox CheckCA 
      Height          =   300
      Index           =   0
      Left            =   9300
      TabIndex        =   7
      Top             =   3750
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox TxtA1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Left            =   1500
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3400
      Width           =   7500
   End
   Begin VB.OptionButton OptCA 
      Height          =   300
      Index           =   4
      Left            =   9300
      TabIndex        =   16
      Top             =   8350
      Width           =   300
   End
   Begin VB.OptionButton OptCA 
      Height          =   300
      Index           =   3
      Left            =   9300
      TabIndex        =   15
      Top             =   7200
      Width           =   300
   End
   Begin VB.OptionButton OptCA 
      Height          =   300
      Index           =   2
      Left            =   9300
      TabIndex        =   14
      Top             =   6050
      Width           =   300
   End
   Begin VB.OptionButton OptCA 
      Height          =   300
      Index           =   1
      Left            =   9300
      TabIndex        =   13
      Top             =   4900
      Width           =   300
   End
   Begin VB.OptionButton OptCA 
      Height          =   300
      Index           =   0
      Left            =   9300
      TabIndex        =   12
      Top             =   3750
      Width           =   300
   End
   Begin MSMask.MaskEdBox TxtScore 
      Height          =   360
      Left            =   1500
      TabIndex        =   17
      Top             =   9200
      Width           =   495
      _ExtentX        =   873
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
   Begin VB.CommandButton CmdSave 
      Height          =   660
      Left            =   5200
      Picture         =   "FrmTestEdit4.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   19
      ToolTipText     =   "Save"
      Top             =   9500
      Width           =   1320
   End
   Begin VB.CommandButton CmdClose 
      Height          =   660
      Left            =   3200
      Picture         =   "FrmTestEdit4.frx":030A
      Style           =   1  'Grafisch
      TabIndex        =   18
      ToolTipText     =   "Close"
      Top             =   9500
      Width           =   1320
   End
   Begin VB.TextBox TxtQuest 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   1500
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1100
      Width           =   7500
   End
   Begin VB.ComboBox CmbTyp 
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
      ItemData        =   "FrmTestEdit4.frx":0614
      Left            =   2000
      List            =   "FrmTestEdit4.frx":0624
      Style           =   2  'Dropdown-Liste
      TabIndex        =   20
      Top             =   400
      Width           =   2800
   End
   Begin VB.Label Label4 
      Caption         =   "Answer 1:"
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
      Left            =   405
      TabIndex        =   24
      Top             =   3750
      Width           =   900
   End
   Begin VB.Label Label9 
      Caption         =   "Points:"
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
      Top             =   9250
      Width           =   1005
   End
   Begin VB.Label Label3 
      Caption         =   "Question:"
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
      TabIndex        =   22
      Top             =   1900
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "Question type:"
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
      TabIndex        =   21
      Top             =   480
      Width           =   1400
   End
End
Attribute VB_Name = "FrmTestEdit4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private y As Byte, rstmp As String
Private Sub Form_Load()
 CmbTyp.ListIndex = 0
 TxtScore.Text = 5
End Sub
Private Sub CheckCA_Click(Index As Integer)
 y = 0
 If CheckCA(0).Value = 1 And TxtA1.Text = "" Then y = y + 1
 If CheckCA(1).Value = 1 And TxtA2.Text = "" Then y = y + 1
 If CheckCA(2).Value = 1 And TxtA3.Text = "" Then y = y + 1
 If CheckCA(3).Value = 1 And TxtA4.Text = "" Then y = y + 1
 If CheckCA(4).Value = 1 And TxtA5.Text = "" Then y = y + 1
 If y <> 0 Then
  MsgBox "If this answer should be correct, enter a valid text", vbExclamation, MsgT
  CheckCA(Index).Value = 0
 End If
End Sub
Private Sub OptCA_Click(Index As Integer)
 If CmbTyp.ListIndex = 0 Then
  y = 0
  If OptCA(0).Value = True And TxtA1.Text = "" Then y = y + 1
  If OptCA(1).Value = True And TxtA2.Text = "" Then y = y + 1
  If OptCA(2).Value = True And TxtA3.Text = "" Then y = y + 1
  If OptCA(3).Value = True And TxtA4.Text = "" Then y = y + 1
  If OptCA(4).Value = True And TxtA5.Text = "" Then y = y + 1
  If y <> 0 Then
   MsgBox "If this answer should be correct, enter a valid text", vbExclamation, MsgT
   OptCA(Index).Value = False
  End If
 End If
End Sub
Private Sub TxtScore_GotFocus()
 TxtScore.SelStart = 0
End Sub
Private Sub TxtScore_Change()
 If TxtScore.Text = "" Or Val(TxtScore.Text) = 0 Then
  MsgBox "The points should be > 0 ...", vbExclamation, MsgT
  TxtScore.Text = 5
 End If
End Sub
Private Sub CmbTyp_Click()
 ResetForm
 Select Case CmbTyp.ListIndex
  Case 0
   Frame1.Visible = True
   Frame2.Visible = True
   For z = 0 To 4
    CheckCA(z).Visible = False
    OptCA(z).Visible = True
   Next
  Case 1
   Frame1.Visible = True
   Frame2.Visible = True
   For z = 0 To 4
    OptCA(z).Visible = False
    CheckCA(z).Visible = True
   Next
  Case 2
   Frame1.Visible = True
   Frame2.Visible = False
   For z = 0 To 4
    CheckCA(z).Visible = False
   Next
   OptCA(0).Visible = True
   OptCA(1).Visible = True
   For z = 2 To 4
    OptCA(z).Visible = False
   Next
  Case 3
   Frame1.Visible = False
   Frame2.Visible = False
   TxtA1.Visible = False
   TxtA1E.Visible = True
   For z = 0 To 4
    OptCA(z).Visible = False
    CheckCA(z).Visible = False
   Next
 End Select
 If CmbTyp.ListIndex = 2 Then
  TxtA1.Enabled = False
  TxtA2.Enabled = False
  TxtA1.Text = "Yes"
  TxtA2.Text = "No"
 Else
  TxtA1.Enabled = True
  TxtA2.Enabled = True
 End If
End Sub
Private Sub CmdSave_Click()
 z = 0
 If TxtQuest.Text = "" Then z = z + 1
 If CmbTyp.ListIndex = 3 Then
  If TxtA1E.Text = "" Then z = z + 1
 Else
  If TxtA1.Text = "" Then z = z + 1
  Select Case CmbTyp.ListIndex
   Case 0
    If TxtA2.Text = "" Then z = z + 1
    If TxtA3.Text = "" Then z = z + 1
    If OptCA(0).Value = False And OptCA(1).Value = False And OptCA(2).Value = False And OptCA(3).Value = False And OptCA(4).Value = False Then z = z + 1
   Case 1
    If TxtA2.Text = "" Then z = z + 1
    If TxtA3.Text = "" Then z = z + 1
    If CheckCA(0).Value = 0 And CheckCA(1).Value = 0 And CheckCA(2).Value = 0 And CheckCA(3).Value = 0 And CheckCA(4).Value = 0 Then z = z + 1
   Case 2
    If TxtA2.Text = "" Then z = z + 1
    If OptCA(0).Value = False And OptCA(1).Value = False Then z = z + 1
  End Select
 End If
 If z = 0 Then
  tmpval = ""
  rstmp = ""
  Select Case CmbTyp.ListIndex
   Case 0
    y = 0
    For z = 0 To 4
     If OptCA(z).Value = True Then y = z + 1
    Next
    tmpval = "INSERT INTO dbsv_test_cat (tcorder,tctid,tctyp,tcquest,tcans1,tcans2,tcans3,tcans4,tcans5,tcanswer,tcpoints) VALUES ('999','" & CmdSave.Tag & "','0','" & TxtQuest.Text & "','" & TxtA1.Text & "','" & TxtA2.Text & "','" & TxtA3.Text & "','" & IIf(TxtA4.Text = "", "0", TxtA4.Text) & "','" & IIf(TxtA5.Text = "", "0", TxtA5.Text) & "','" & y & "','" & TxtScore.Text & "');"
   Case 1
    For z = 0 To 4
     rstmp = rstmp & IIf(CheckCA(z).Value = 1, 1, 0)
    Next
    tmpval = "INSERT INTO dbsv_test_cat (tcorder,tctid,tctyp,tcquest,tcans1,tcans2,tcans3,tcans4,tcans5,tcanswer,tcpoints) VALUES ('999','" & CmdSave.Tag & "','1','" & TxtQuest.Text & "','" & TxtA1.Text & "','" & TxtA2.Text & "','" & TxtA3.Text & "','" & IIf(TxtA4.Text = "", "0", TxtA4.Text) & "','" & IIf(TxtA5.Text = "", "0", TxtA5.Text) & "','" & rstmp & "','" & TxtScore.Text & "');"
   Case 2
    tmpval = "INSERT INTO dbsv_test_cat (tcorder,tctid,tctyp,tcquest,tcans1,tcans2,tcans3,tcans4,tcans5,tcanswer,tcpoints) VALUES ('999','" & CmdSave.Tag & "','2','" & TxtQuest.Text & "','" & TxtA1.Text & "','" & TxtA2.Text & "','0','0','0','" & IIf(OptCA(0).Value = True, 1, 2) & "','" & TxtScore.Text & "');"
   Case 3
    tmpval = "INSERT INTO dbsv_test_cat (tcorder,tctid,tctyp,tcquest,tcans1,tcans2,tcans3,tcans4,tcans5,tcanswer,tcpoints) VALUES ('999','" & CmdSave.Tag & "','3','" & TxtQuest.Text & "','" & TxtA1.Text & "','0','0','0','0','1','" & TxtScore.Text & "');"
  End Select
  dbcon.Execute tmpval
  rscon.Open "SELECT last_insert_id();", dbcon, adOpenDynamic, adLockOptimistic
  tmpval = rscon.Fields("last_insert_id()").Value
  rscon.Close
  With FrmTestEdit2.List1
   .AddItem Left$(TxtQuest.Text, 30)
   .ItemData(.NewIndex) = Val(tmpval)
   .Refresh
  End With
  If CmdClose.Tag = 0 Then dbcon.Execute "UPDATE dbsv_test_setting SET tsactive='1' WHERE tsid='" & CmdSave.Tag & "';"
  tmpval = ""
  rstmp = ""
  Unload Me
 Else
  MsgBox "Some fields are missing." & vbCrLf & "Check following:" & vbCrLf & "* Entered a valid question" & vbCrLf & "* Entered possible answers (min. 3 when Multiple Choice)" & vbCrLf & "* Selected correct answer(s)", vbExclamation, MsgT
 End If
End Sub
Private Sub CmdClose_Click()
 Unload Me
End Sub
Private Sub ResetForm()
 TxtQuest.Text = ""
 TxtA1.Text = ""
 TxtA2.Text = ""
 TxtA3.Text = ""
 TxtA4.Text = ""
 TxtA5.Text = ""
 TxtA1E.Visible = False
 TxtA1.Visible = True
 TxtScore.Text = 5
 For z = 0 To 4
  OptCA(z).Value = False
  CheckCA(z).Value = 0
 Next
End Sub
