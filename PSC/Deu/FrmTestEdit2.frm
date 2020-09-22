VERSION 5.00
Begin VB.Form FrmTestEdit2 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "TestEditor - Aufgaben"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton CmdSort 
      Height          =   600
      Left            =   5300
      Picture         =   "FrmTestEdit2.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Aufgaben sortieren"
      Top             =   3400
      Width           =   1050
   End
   Begin VB.CommandButton CmdDel 
      DisabledPicture =   "FrmTestEdit2.frx":014A
      Height          =   600
      Left            =   5300
      Picture         =   "FrmTestEdit2.frx":051A
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Aufgabe entfernen"
      Top             =   2500
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton CmdAdd 
      Height          =   600
      Left            =   5300
      Picture         =   "FrmTestEdit2.frx":095C
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Aufgabe hinzufügen"
      Top             =   1600
      Width           =   1050
   End
   Begin VB.CommandButton CmdClose 
      Height          =   600
      Left            =   5300
      Picture         =   "FrmTestEdit2.frx":1626
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Verlassen"
      Top             =   4400
      Width           =   1050
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
      Height          =   6060
      Left            =   450
      TabIndex        =   4
      Top             =   1200
      Width           =   4500
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   900
      TabIndex        =   6
      Top             =   300
      Width           =   5000
   End
   Begin VB.Label Label3 
      Caption         =   "Test-Aufgaben:"
      Height          =   300
      Left            =   500
      TabIndex        =   5
      Top             =   800
      Width           =   1200
   End
   Begin VB.Menu mnufile 
      Caption         =   "Datei"
      Begin VB.Menu mnuclose 
         Caption         =   "Schliessen"
      End
   End
   Begin VB.Menu mnutask 
      Caption         =   "Aufgaben"
      Begin VB.Menu mnuadd 
         Caption         =   "Hinzufügen"
      End
      Begin VB.Menu mnudel 
         Caption         =   "Entfernen"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusort 
         Caption         =   "Sortieren"
      End
   End
End
Attribute VB_Name = "FrmTestEdit2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub mnuclose_Click()
 Unload Me
End Sub
Private Sub mnuadd_Click()
 CmdAdd_Click
End Sub
Private Sub mnudel_Click()
 CmdDel_Click
End Sub
Private Sub mnusort_Click()
 CmdSort_Click
End Sub
Private Sub CmdAdd_Click()
 Load FrmTestEdit4
 With FrmTestEdit4
  .CmdSave.Tag = CmdClose.Tag
  .CmdClose.Tag = CmdDel.Tag
  .Show vbModal
 End With
 SetBtT False
End Sub
Private Sub CmdDel_Click()
 If vbYes = MsgBox("Diese Aufgabe aus dem Test entfernen?", vbExclamation + vbYesNo, MsgT) Then
  dbcon.Execute "DELETE FROM dbsv_test_cat WHERE tcid='" & IIf(List1.ListIndex = 0, "1", List1.ItemData(List1.ListIndex)) & "' AND tctid='" & CmdClose.Tag & "';"
  List1.RemoveItem List1.ListIndex
  List1.Refresh
  SetBtT False
  If List1.ListCount = 0 Then dbcon.Execute "UPDATE dbsv_test_setting SET tsactive='0' WHERE tsid='" & CmdClose.Tag & "';"
 End If
End Sub
Private Sub CmdSort_Click()
 If List1.ListCount > 1 Then
  Load FrmTestEdit3
  For z = 0 To List1.ListCount - 1
   FrmTestEdit3.List1.AddItem List1.List(z)
   FrmTestEdit3.List1.ItemData(FrmTestEdit3.List1.NewIndex) = List1.ItemData(z)
  Next
  Unload Me
  FrmTestEdit3.Show vbModal
 Else
  MsgBox "Zur Sortierung sollte der Test mindestens 2 Aufgaben enthalten", vbExclamation, MsgT
 End If
End Sub
Private Sub CmdClose_Click()
 Unload Me
End Sub
Private Sub List1_Click()
 If List1.ListIndex = -1 Then SetBtT False Else SetBtT True
End Sub
Function LoadFrm(vTest As Integer, vTestName As String)
 CmdClose.Tag = vTest
 rscon.Open "SELECT a.tsactive,b.tcid,b.tctid,b.tcquest FROM dbsv_test_cat AS b JOIN dbsv_test_setting AS a ON b.tctid=a.tsid WHERE b.tctid='" & vTest & "' ORDER BY b.tcorder,b.tcid;", dbcon, adOpenDynamic, adLockOptimistic
 If rscon.EOF = False Then
  CmdDel.Tag = rscon.Fields("tsactive").Value
  Do While rscon.EOF = False
   List1.AddItem Left$(rscon.Fields("tcquest").Value, 30)
   List1.ItemData(List1.NewIndex) = rscon.Fields("tcid").Value
   rscon.MoveNext
  Loop
 Else
  CmdDel.Tag = 0
 End If
 rscon.Close
 Label1.Caption = "Test: " & vTestName
 Me.Show vbModal
End Function
Private Function SetBtT(Mode As Boolean)
 mnudel.Enabled = Mode
 CmdDel.Visible = Mode
End Function
