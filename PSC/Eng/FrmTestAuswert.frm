VERSION 5.00
Begin VB.Form FrmTestAuswert 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Test Results"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton CmdDel 
      Enabled         =   0   'False
      Height          =   660
      Left            =   5880
      Picture         =   "FrmTestAuswert.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Delete Result"
      Top             =   2500
      Width           =   1320
   End
   Begin VB.CommandButton CmdClose 
      Height          =   660
      Left            =   5900
      Picture         =   "FrmTestAuswert.frx":0442
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Close"
      Top             =   3500
      Width           =   1320
   End
   Begin VB.CommandButton CmdDet 
      Enabled         =   0   'False
      Height          =   660
      Left            =   5900
      Picture         =   "FrmTestAuswert.frx":074C
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Details"
      Top             =   1500
      Width           =   1320
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   400
      TabIndex        =   3
      Top             =   900
      Width           =   5200
   End
   Begin VB.Label Label1 
      Caption         =   "Test Results: (Select one for Details)"
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
      TabIndex        =   4
      Top             =   300
      Width           =   3600
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnutest 
      Caption         =   "Test"
      Begin VB.Menu mnudet 
         Caption         =   "Show Details"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnudel 
         Caption         =   "Delete Result"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmTestAuswert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub mnudet_Click()
 CmdDet_Click
End Sub
Private Sub mnudel_Click()
 CmdDel_Click
End Sub
Private Sub mnuclose_Click()
 Unload Me
End Sub
Private Sub List1_Click()
 If List1.ListIndex <> -1 Then SetBt True Else SetBt False
End Sub
Private Sub CmdDet_Click()
 FrmTestAuswert2.LoadFrm List1.ItemData(List1.ListIndex), CmdDet.Tag
 List1.ListIndex = -1
End Sub
Private Sub CmdDel_Click()
 If vbYes = MsgBox("Delete this result?", vbExclamation + vbYesNo, MsgT) Then
  dbcon.Execute "DELETE FROM dbsv_test_result WHERE trid='" & List1.ItemData(List1.ListIndex) & "';"
  List1.RemoveItem List1.ListIndex
 End If
 List1.ListIndex = -1
End Sub
Private Sub CmdClose_Click()
 Unload Me
End Sub
Function LoadFrm(vTest As Integer, vTestName As String)
 CmdDet.Tag = vTestName
 If UsrInf.Role = 2 Then
  rscon.Open "SELECT trid,trdatum FROM dbsv_test_result WHERE truid='" & UsrInf.ID & "' AND trtid='" & vTest & "';", dbcon, adOpenDynamic, adLockOptimistic
  If rscon.EOF = False Then
   Do While rscon.EOF = False
    List1.AddItem "Versuch am: " & rscon.Fields("trdatum").Value
    List1.ItemData(List1.NewIndex) = rscon.Fields("trid").Value
    rscon.MoveNext
   Loop
  Else
   List1.AddItem "No results"
   List1.AddItem "for this test available"
   List1.Enabled = False
  End If
  CmdDel.Visible = False
  mnudel.Visible = False
 Else
  rscon.Open "SELECT a.trid,a.trdatum,b.uname FROM dbsv_test_result AS a JOIN dbsv_adm_user AS b ON a.truid=b.uid WHERE a.trtid='" & vTest & "' ORDER BY a.trtid,a.truid,a.trdatum;", dbcon, adOpenDynamic, adLockOptimistic
  If rscon.EOF = False Then
   Do While rscon.EOF = False
    List1.AddItem rscon.Fields("uname").Value & "  " & rscon.Fields("trdatum").Value
    List1.ItemData(List1.NewIndex) = rscon.Fields("trid").Value
    rscon.MoveNext
   Loop
  Else
   List1.AddItem "No results"
   List1.AddItem "for this test available"
   List1.Enabled = False
  End If
 End If
 rscon.Close
 Me.Show vbModal
End Function
Private Function SetBt(Mode As Boolean)
 mnudet.Enabled = Mode
 CmdDet.Enabled = Mode
 mnudel.Enabled = Mode
 CmdDel.Enabled = Mode
End Function
