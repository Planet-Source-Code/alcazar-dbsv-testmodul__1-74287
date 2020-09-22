VERSION 5.00
Begin VB.Form FrmTestS2 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Run Test"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox CmbQuest 
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
      Left            =   400
      Style           =   2  'Dropdown-Liste
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.CheckBox CheckCA 
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   9400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.OptionButton OptCA 
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   8900
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
      Height          =   400
      Left            =   1400
      MaxLength       =   200
      TabIndex        =   36
      Top             =   3300
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.CommandButton CmdEnd 
      Height          =   660
      Left            =   6200
      Picture         =   "FrmTestS2.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "End Test"
      Top             =   8900
      Width           =   1320
   End
   Begin VB.PictureBox PicTime 
      BorderStyle     =   0  'Kein
      Height          =   600
      Left            =   7200
      Picture         =   "FrmTestS2.frx":0CCA
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Zeit"
      Top             =   350
      Width           =   600
   End
   Begin VB.CommandButton CmdNext 
      Height          =   660
      Left            =   4300
      Picture         =   "FrmTestS2.frx":1169
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Next question"
      Top             =   8900
      Width           =   1320
   End
   Begin VB.CommandButton CmdCancel 
      Height          =   660
      Left            =   2400
      Picture         =   "FrmTestS2.frx":15AB
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Cancel Test"
      Top             =   8900
      Width           =   1320
   End
   Begin VB.OptionButton OptCA 
      Height          =   300
      Index           =   1
      Left            =   9350
      TabIndex        =   3
      Top             =   3600
      Width           =   300
   End
   Begin VB.OptionButton OptCA 
      Height          =   300
      Index           =   2
      Left            =   9350
      TabIndex        =   4
      Top             =   4600
      Width           =   300
   End
   Begin VB.OptionButton OptCA 
      Height          =   300
      Index           =   3
      Left            =   9350
      TabIndex        =   5
      Top             =   5700
      Width           =   300
   End
   Begin VB.OptionButton OptCA 
      Height          =   300
      Index           =   4
      Left            =   9350
      TabIndex        =   6
      Top             =   6750
      Width           =   300
   End
   Begin VB.OptionButton OptCA 
      Height          =   300
      Index           =   5
      Left            =   9350
      TabIndex        =   7
      Top             =   7800
      Width           =   300
   End
   Begin VB.CheckBox CheckCA 
      Height          =   300
      Index           =   1
      Left            =   9350
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CheckBox CheckCA 
      Height          =   255
      Index           =   2
      Left            =   9350
      TabIndex        =   9
      Top             =   4600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox CheckCA 
      Height          =   300
      Index           =   5
      Left            =   9350
      TabIndex        =   12
      Top             =   7800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CheckBox CheckCA 
      Height          =   300
      Index           =   4
      Left            =   9350
      TabIndex        =   11
      Top             =   6750
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CheckBox CheckCA 
      Height          =   300
      Index           =   3
      Left            =   9350
      TabIndex        =   10
      Top             =   5700
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Kein
      Height          =   1100
      Left            =   350
      TabIndex        =   26
      Top             =   4300
      Width           =   8900
      Begin VB.Label LblA2 
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
         Left            =   1100
         TabIndex        =   32
         Top             =   50
         Width           =   7700
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
         Top             =   50
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'Kein
      Height          =   3200
      Left            =   350
      TabIndex        =   22
      Top             =   5400
      Width           =   8900
      Begin VB.Label LblA5 
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
         Left            =   1100
         TabIndex        =   35
         Top             =   2100
         Width           =   7700
      End
      Begin VB.Label LblA4 
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
         Left            =   1100
         TabIndex        =   34
         Top             =   1100
         Width           =   7700
      End
      Begin VB.Label LblA3 
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
         Left            =   1100
         TabIndex        =   33
         Top             =   50
         Width           =   7700
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
         TabIndex        =   25
         Top             =   50
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
         TabIndex        =   24
         Top             =   1100
         Width           =   900
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
         TabIndex        =   23
         Top             =   2100
         Width           =   900
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6500
      Top             =   400
   End
   Begin VB.Frame FrameTime 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   7900
      TabIndex        =   14
      Top             =   250
      Width           =   1800
      Begin VB.Label LblSec 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1300
         TabIndex        =   19
         Top             =   250
         Width           =   300
      End
      Begin VB.Label Label2 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1100
         TabIndex        =   18
         Top             =   250
         Width           =   150
      End
      Begin VB.Label LblMin 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   750
         TabIndex        =   17
         Top             =   250
         Width           =   300
      End
      Begin VB.Label Label8 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   550
         TabIndex        =   16
         Top             =   250
         Width           =   150
      End
      Begin VB.Label LblStd 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   200
         TabIndex        =   15
         Top             =   250
         Width           =   300
      End
   End
   Begin VB.Label LblA1 
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
      Left            =   1400
      TabIndex        =   31
      Top             =   3300
      Width           =   7700
   End
   Begin VB.Label LblCurA 
      Alignment       =   2  'Zentriert
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
      Left            =   500
      TabIndex        =   30
      Top             =   2400
      Width           =   500
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
      Left            =   400
      TabIndex        =   28
      Top             =   3400
      Width           =   900
   End
   Begin VB.Label LblTest 
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
      Top             =   400
      Width           =   4900
   End
   Begin VB.Label LblQuest 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1900
      Left            =   1400
      TabIndex        =   20
      Top             =   1200
      Width           =   7700
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
      TabIndex        =   13
      Top             =   1950
      Width           =   900
   End
End
Attribute VB_Name = "FrmTestS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private answer() As String, useranswer() As String, points() As Byte, tpass As Byte
Private gaufg As Integer, ca As Integer, uscore As Long, score_max As Long
Private Const MsgText As String = "Please enter a valid answer for that question"
Private Sub Timer1_Timer()
 If Val(LblStd.Caption) = 0 And Val(LblMin.Caption) = 0 And Val(LblSec.Caption) = 0 Then
  Timer1.Enabled = False
  MsgBox "Time is over, Test will be ended!", vbExclamation, MsgT
  EndTest 1
  Exit Sub
 End If
 If Val(LblStd.Caption) = 0 And Val(LblMin.Caption) = 5 And Val(LblSec.Caption) = 0 And LblMin.ForeColor <> vbRed Then
  LblStd.ForeColor = vbRed
  LblMin.ForeColor = vbRed
  LblSec.ForeColor = vbRed
 End If
 If Val(LblSec.Caption) = 0 Then
  If Val(LblMin.Caption) = 0 And Val(LblStd.Caption) > 0 Then
   LblStd.Caption = LblStd.Caption - 1
   If Val(LblStd.Caption) < 10 Then LblStd.Caption = "0" & LblStd.Caption
   LblMin.Caption = "59"
   LblSec.Caption = "59"
   Exit Sub
  End If
  LblMin.Caption = LblMin.Caption - 1
  If Val(LblMin.Caption) < 10 Then LblMin.Caption = "0" & LblMin.Caption
  LblSec.Caption = "59"
 Else
  LblSec.Caption = LblSec.Caption - 1
  If Val(LblSec.Caption) < 10 Then LblSec.Caption = "0" & LblSec.Caption
 End If
End Sub
Private Sub CmdCancel_Click()
 If vbYes = MsgBox("Really cancel Test?" & vbCrLf & "Test will be marked as ""not passed""", vbExclamation + vbYesNo, MsgT) Then
  Timer1.Enabled = False
  If rscon.State = adStateOpen Then rscon.Close
  If CmdCancel.Tag = False Then
   dbcon.Execute "INSERT INTO dbsv_test_result (truid,trtid,trfid,trdatum,trquestg,trscore,trscore_max,trpass) VALUES ('" & UsrInf.ID & "','" & CmdEnd.Tag & "','" & CmdNext.Tag & "','" & CStr(Now()) & "','" & gaufg & "','0','0','0');"
   Log_Entry "D", "Canceld Test: " & Mid$(LblTest.Caption, 8), UsrInf.ID
  End If
  Unload Me
 End If
End Sub
Private Sub CmdNext_Click()
 If LblCurA.Caption = gaufg Then
  MsgBox """Next Question"" not possible" & vbCrLf & "Its the last question", vbExclamation, MsgT
 Else
  If CheckAns = False Then Exit Sub
  LblCurA.Caption = LblCurA.Caption + 1
  If Val(LblCurA.Caption) = gaufg Then CmdEnd.Visible = True
  ShowAufg 1
  If useranswer(ca + 1) <> "" Then ShowUAns
 End If
End Sub
Private Sub CmdEnd_Click()
 If vbYes = MsgBox("Do you want to end the test and have it scored?" & vbCrLf & "Not answered questions are treated as ""wrong answer""!", vbExclamation + vbYesNo, MsgT) Then EndTest 0
End Sub
Private Sub CmbQuest_Click()
 If CmbQuest.ListIndex <> 0 Then
  Debug.Print CmbQuest.ListIndex
 End If
End Sub
Private Sub ShowAufg(Mode As Integer)
 If Mode = 1 Then rscon.MoveNext
 LblQuest.Caption = rscon.Fields("tcquest").Value
 CmdNext.Tag = rscon.Fields("tctyp").Value
 TxtA1.Visible = False
 LblA1.Visible = True
 For z = 1 To 5
  OptCA(z).Enabled = True
  OptCA(z).Value = False
  CheckCA(z).Enabled = True
  CheckCA(z).Value = 0
 Next
 Select Case CmdNext.Tag
  Case 0
   Frame1.Visible = True
   Frame2.Visible = True
   For z = 1 To 5
    CheckCA(z).Visible = False
    OptCA(z).Visible = True
   Next
   LblA1.Caption = rscon.Fields("tcans1").Value
   LblA2.Caption = rscon.Fields("tcans2").Value
   LblA3.Caption = rscon.Fields("tcans3").Value
   If rscon.Fields("tcans4").Value = "0" Then
    LblA4.Caption = ""
    OptCA(4).Enabled = False
   Else
    LblA4.Caption = rscon.Fields("tcans4").Value
   End If
   If rscon.Fields("tcans5").Value = "0" Then
    LblA5.Caption = ""
    OptCA(5).Enabled = False
   Else
    LblA5.Caption = rscon.Fields("tcans5").Value
   End If
  Case 1
   Frame1.Visible = True
   Frame2.Visible = True
   For z = 1 To 5
    OptCA(z).Visible = False
    CheckCA(z).Visible = True
   Next
   LblA1.Caption = rscon.Fields("tcans1").Value
   LblA2.Caption = rscon.Fields("tcans2").Value
   LblA3.Caption = rscon.Fields("tcans3").Value
   If rscon.Fields("tcans4").Value = "0" Then
    LblA4.Caption = ""
    CheckCA(4).Enabled = False
   Else
    LblA4.Caption = rscon.Fields("tcans4").Value
   End If
   If rscon.Fields("tcans5").Value = "0" Then
    LblA5.Caption = ""
    CheckCA(5).Enabled = False
   Else
    LblA5.Caption = rscon.Fields("tcans4").Value
   End If
  Case 2
   Frame1.Visible = True
   Frame2.Visible = False
   For z = 1 To 5
    CheckCA(z).Visible = False
   Next
   OptCA(1).Visible = True
   OptCA(2).Visible = True
   For z = 3 To 5
    OptCA(z).Visible = False
   Next
   LblA1.Caption = rscon.Fields("tcans1").Value
   LblA2.Caption = rscon.Fields("tcans2").Value
  Case 3
   Frame1.Visible = False
   Frame2.Visible = False
   For z = 1 To 5
    OptCA(z).Visible = False
    CheckCA(z).Visible = False
   Next
   LblA1.Visible = False
   TxtA1.Visible = True
   TxtA1.SetFocus
 End Select
 answer(Val(LblCurA.Caption)) = rscon.Fields("tcanswer").Value
 points(Val(LblCurA.Caption)) = rscon.Fields("tcpoints").Value
End Sub
Private Sub ShowUAns()
 Select Case CmdNext.Tag
  Case 0, 2
   OptCA(useranswer(ca + 1)).Value = True
  Case 1
   For z = 1 To 5
    CheckCA(z).Value = Mid$(useranswer(ca + 1), z, 1)
   Next
  Case 3
   TxtA1.Text = useranswer(ca + 1)
 End Select
End Sub
Private Sub EndTest(Mode As Integer)
 Timer1.Enabled = False
 If rscon.State = adStateOpen Then rscon.Close
 If CmdCancel.Tag = False Then
  uscore = 0
  tmpval = ""
  For z = 1 To gaufg
   score_max = score_max + points(z)
   If useranswer(z) = answer(z) Then uscore = uscore + points(z)
  Next
  If Mode = 1 And TestInf.TimeExp = 1 Then
   tpass = 0
  Else
   tpass = IIf(((uscore * 100) / score_max) >= TestInf.ScoreToPass, 1, 0)
  End If
  dbcon.Execute "INSERT INTO dbsv_test_result (truid,trtid,trfid,trdatum,trquestg,trscore,trscore_max,trpass) VALUES ('" & UsrInf.ID & "','" & CmdEnd.Tag & "','" & CmdNext.Tag & "','" & CStr(Now()) & "','" & gaufg & "','" & uscore & "','" & score_max & "','" & tpass & "');"
  rscon.Open "SELECT last_insert_id();", dbcon, adOpenDynamic, adLockOptimistic
  tmpval = rscon.Fields("last_insert_id()").Value
  rscon.Close
  Log_Entry "D", "Test Taken: " & Mid$(LblTest.Caption, 8), UsrInf.ID
  Me.Hide
  FrmTestAuswert2.LoadFrm Val(tmpval), Mid$(LblTest.Caption, 8)
 Else
  Unload Me
 End If
End Sub
Private Function CheckAns() As Boolean
 ca = Val(LblCurA.Caption)
 tmpval = ""
 CheckAns = False
 Select Case CmdNext.Tag
  Case 0
   If OptCA(1).Value = False And OptCA(2).Value = False And OptCA(3).Value = False And OptCA(4).Value = False And OptCA(5).Value = False Then
    MsgBox MsgText, vbExclamation, MsgT
    Exit Function
   Else
    For z = 1 To 5
     If OptCA(z).Value = True Then useranswer(ca) = z
    Next
   End If
  Case 1
   If CheckCA(1).Value = 0 And CheckCA(2).Value = 0 And CheckCA(3).Value = 0 And CheckCA(4).Value = 0 And CheckCA(5).Value = 0 Then
    MsgBox MsgText, vbExclamation, MsgT
    Exit Function
   Else
    For z = 1 To 5
     tmpval = tmpval & IIf(CheckCA(z).Value = 1, 1, 0)
    Next
    useranswer(ca) = tmpval
   End If
  Case 2
   If OptCA(1).Value = False And OptCA(2).Value = False Then
    MsgBox MsgText, vbExclamation, MsgT
    Exit Function
   Else
    useranswer(ca) = IIf(OptCA(1).Value = True, "1", "2")
   End If
  Case 3
   If TxtA1.Text = "" Then
    MsgBox MsgText, vbExclamation, MsgT
    Exit Function
   Else
    useranswer(ca) = TxtA1.Text
   End If
 End Select
 tmpval = ""
 CheckAns = True
End Function
Function LoadFrm(vTest As Integer, vTestName As String, vFach As Integer, vMode As Boolean)
 Dim introtxt As String, mlimit As Integer, tshow As Boolean

 CmdEnd.Tag = vTest
 CmdNext.Tag = vFach
 CmdCancel.Tag = vMode
 LblTest.Caption = "Test:  " & vTestName
 rscon.Open "SELECT tsintro,tstimelimit,tstime_exp,tsmultilimit,tsshowq,tsscore_pass FROM dbsv_test_setting WHERE tsid='" & vTest & "';", dbcon, adOpenDynamic, adLockOptimistic
 With TestInf
  .ScoreToPass = rscon.Fields("tsscore_pass").Value
  .TimeExp = rscon.Fields("tstime_exp").Value
  .TimeL = rscon.Fields("tstimelimit").Value
 End With
 introtxt = rscon.Fields("tsintro").Value
 mlimit = rscon.Fields("tsmultilimit").Value
 tshow = CBool(rscon.Fields("tsshowq").Value)
 rscon.Close
 If TestInf.TimeL <> 0 Then
  Select Case TestInf.TimeL
   Case Is < 60
    LblMin.Caption = TestInf.TimeL
   Case 60 To 119
    LblStd.Caption = "01"
    LblMin.Caption = TestInf.TimeL - 60
   Case 120 To 179
    LblStd.Caption = "02"
    LblMin.Caption = TestInf.TimeL - 120
   Case 180 To 239
    LblStd.Caption = "03"
    LblMin.Caption = TestInf.TimeL - 180
   Case 240
    LblStd.Caption = "04"
  End Select
  If Val(LblMin.Caption) < 10 And LblStd.Caption <> "04" Then LblMin.Caption = "0" & LblMin.Caption
 End If
 rscon.Open "SELECT COUNT(*) AS anzahl FROM dbsv_test_cat WHERE tctid='" & vTest & "';", dbcon, adOpenDynamic, adLockOptimistic
 gaufg = rscon.Fields("anzahl").Value
 rscon.Close
 With FrmTestIntro
  .Label1.Caption = introtxt
  .Label2.Caption = "This Test contains " & gaufg & " Questions and can be taken " & mlimit & "x ."
  If TestInf.TimeL = 0 Then
   .Label3.Caption = "There is no timelimit in this test."
  Else
   .Label3.Caption = "This Test has a timelimit of " & TestInf.TimeL & " minutes."
  End If
  .Show vbModal
 End With
 Me.Caption = Me.Caption & " (" & gaufg & " Aufgaben)"
 LblCurA.Caption = "1"
 If tshow = True Then
  LblCurA.Visible = False
  CmdNext.Visible = False
  CmbQuest.Visible = True
  For z = 1 To gaufg
   CmbQuest.AddItem z
  Next
  CmbQuest.ListIndex = 0
 End If
 ReDim answer(gaufg)
 ReDim useranswer(gaufg)
 ReDim points(gaufg)
 For z = 1 To UBound(answer)
  answer(z) = ""
  useranswer(z) = ""
  points(z) = 0
 Next
 If TestInf.TimeL <> 0 Then Timer1.Enabled = True
 rscon.Open "SELECT * FROM dbsv_test_cat WHERE tctid='" & vTest & "';", dbcon, adOpenDynamic, adLockOptimistic
 ShowAufg 0
 Me.Show vbModal
End Function
