VERSION 5.00
Begin VB.Form FrmTestAuswert2 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "TestErgebnisse - Details"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton CmdPrint 
      Caption         =   "Drucken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1400
      TabIndex        =   1
      Top             =   4600
      Width           =   1300
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Schliessen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3400
      Picture         =   "FrmTestAuswert2.frx":0000
      TabIndex        =   0
      ToolTipText     =   "Schliessen"
      Top             =   4600
      Width           =   1300
   End
   Begin VB.PictureBox PicNoPass 
      BorderStyle     =   0  'Kein
      Height          =   400
      Left            =   1600
      Picture         =   "FrmTestAuswert2.frx":030A
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Nein"
      Top             =   3900
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.PictureBox PicPass 
      BorderStyle     =   0  'Kein
      Height          =   400
      Left            =   1600
      Picture         =   "FrmTestAuswert2.frx":0614
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Ja"
      Top             =   3900
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.Label LblNeededScoreP 
      Caption         =   "( 0 % )"
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
      Left            =   2600
      TabIndex        =   20
      Top             =   3400
      Width           =   900
   End
   Begin VB.Label LblScoreP 
      Caption         =   "( 0 % )"
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
      Left            =   2600
      TabIndex        =   19
      Top             =   2900
      Width           =   900
   End
   Begin VB.Label LblUser 
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
      Left            =   1200
      TabIndex        =   18
      Top             =   400
      Width           =   4600
   End
   Begin VB.Label Label5 
      Caption         =   "Benutzer:"
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
      TabIndex        =   17
      Top             =   400
      Width           =   900
   End
   Begin VB.Label LblNeededScore 
      Caption         =   "0"
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
      Left            =   2100
      TabIndex        =   16
      Top             =   3400
      Width           =   500
   End
   Begin VB.Label Label9 
      Caption         =   "Punktzahl benötigt:"
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
      TabIndex        =   15
      Top             =   3400
      Width           =   1700
   End
   Begin VB.Label LblFach 
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
      Left            =   900
      TabIndex        =   14
      Top             =   1400
      Width           =   4600
   End
   Begin VB.Label Label1 
      Caption         =   "Fach:"
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
      TabIndex        =   13
      Top             =   1400
      Width           =   600
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
      Height          =   300
      Left            =   1800
      TabIndex        =   12
      Top             =   2400
      Width           =   600
   End
   Begin VB.Label LblScore 
      Caption         =   "0"
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
      Left            =   2100
      TabIndex        =   11
      Top             =   2900
      Width           =   500
   End
   Begin VB.Label LblDatum 
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
      Left            =   1000
      TabIndex        =   10
      Top             =   1900
      Width           =   1900
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
      Left            =   850
      TabIndex        =   9
      Top             =   900
      Width           =   4600
   End
   Begin VB.Label Label8 
      Caption         =   "Bestanden?"
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
      Top             =   4000
      Width           =   1200
   End
   Begin VB.Label Label7 
      Caption         =   "Punktzahl erreicht:"
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
      TabIndex        =   5
      Top             =   2900
      Width           =   1700
   End
   Begin VB.Label Label4 
      Caption         =   "Anzahl Fragen:"
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
      TabIndex        =   4
      Top             =   2400
      Width           =   1400
   End
   Begin VB.Label Label3 
      Caption         =   "Datum:"
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
      TabIndex        =   3
      Top             =   1900
      Width           =   700
   End
   Begin VB.Label Label2 
      Caption         =   "Test:"
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
      TabIndex        =   2
      Top             =   900
      Width           =   600
   End
End
Attribute VB_Name = "FrmTestAuswert2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdPrint_Click()
 ' Username als "Header" setzen
 MsgBox "Noch net fertich...", vbInformation, MsgT
End Sub
Private Sub CmdClose_Click()
 Unload Me
End Sub
Function LoadFrm(vTest As Integer, vTestName As String)
 tmpval = ""
 If FrmTestS2.Visible = False Then Unload FrmTestS2
 LblTest.Caption = vTestName
 rscon.Open "SELECT a.trdatum,a.trquestg,a.trscore,a.trscore_max,a.trpass,b.name,c.tsscore_pass,d.uname FROM dbsv_test_result AS a JOIN dbsv_main_fach AS b ON a.trfid=b.fid JOIN dbsv_test_setting AS c ON a.trtid=c.tsid JOIN dbsv_adm_user AS d ON a.truid=d.uid WHERE a.trid='" & vTest & "';", dbcon, adOpenDynamic, adLockOptimistic
 If rscon.EOF = False Then
  LblUser.Caption = rscon.Fields("uname").Value
  LblFach.Caption = rscon.Fields("name").Value
  LblDatum.Caption = rscon.Fields("trdatum").Value
  LblQuest.Caption = rscon.Fields("trquestg").Value
  If rscon.Fields("trscore").Value <> 0 Then
   LblScore.Caption = rscon.Fields("trscore").Value
   LblScoreP.Caption = "( " & (rscon.Fields("trscore").Value * 100) / rscon.Fields("trscore_max").Value & " % )"
  End If
  If rscon.Fields("trscore_max").Value <> 0 Then
   LblNeededScore.Caption = (rscon.Fields("tsscore_pass").Value * rscon.Fields("trscore_max").Value) / 100
   LblNeededScoreP.Caption = "( " & rscon.Fields("tsscore_pass").Value & " % )"
  End If
  If rscon.Fields("trpass").Value = 1 Then PicPass.Visible = True Else PicNoPass.Visible = True
 End If
 rscon.Close
 Me.Show vbModal
End Function
