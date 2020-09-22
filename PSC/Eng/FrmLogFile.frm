VERSION 5.00
Begin VB.Form FrmLogFile 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "LogFile - logged Events"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "FrmLogFile.frx":0000
      Left            =   5900
      List            =   "FrmLogFile.frx":0010
      Style           =   2  'Dropdown-Liste
      TabIndex        =   5
      Top             =   200
      Width           =   1600
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5500
      Left            =   200
      TabIndex        =   2
      Top             =   800
      Width           =   12200
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
         ItemData        =   "FrmLogFile.frx":0034
         Left            =   100
         List            =   "FrmLogFile.frx":0036
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   12000
      End
      Begin VB.Label Label1 
         Caption         =   $"FrmLogFile.frx":0038
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
         Left            =   100
         TabIndex        =   4
         Top             =   200
         Width           =   9700
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   660
      Left            =   6600
      Picture         =   "FrmLogFile.frx":00E2
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   6650
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Height          =   660
      Left            =   4800
      Picture         =   "FrmLogFile.frx":03EC
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Options"
      Top             =   6650
      Width           =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "Filter:"
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
      Left            =   5100
      TabIndex        =   6
      Top             =   250
      Width           =   600
   End
   Begin VB.Menu mnu 
      Caption         =   "Menue"
      Visible         =   0   'False
      Begin VB.Menu mnusavelog 
         Caption         =   "Save Logfile"
      End
      Begin VB.Menu mnudellog 
         Caption         =   "Clear Logfile"
      End
   End
End
Attribute VB_Name = "FrmLogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_SETTABSTOPS = &H192, MaxTab = 3
Private LTab(1 To MaxTab) As Long, fnr As Integer
Private Sub Form_Load()
 If UsrInf.Role <> 1 Then Command2.Visible = False
 LTab(1) = 240
 LTab(2) = 278
 LTab(3) = 340
 Combo1.ListIndex = 0
End Sub
Private Sub mnusavelog_Click()
 If rscon.State = adStateOpen Then rscon.Close
 SaveLog
 MsgBox "LogFile saved", vbInformation, MsgT
End Sub
Private Sub mnudellog_Click()
 If vbYes = MsgBox("Clear LogFile?", vbExclamation + vbYesNo, MsgT) Then
  If rscon.State = adStateOpen Then rscon.Close
  If vbYes = MsgBox("Save LogFile before clearing?", vbExclamation + vbYesNo, MsgT) Then SaveLog
  dbcon.Execute "DELETE FROM dbsv_adm_logfile;"
  Log_Entry "S", "Cleared LogFile", UsrInf.ID
  Unload Me
 End If
End Sub
Private Sub Combo1_Click()
 If rscon.State = adStateOpen Then rscon.Close
 ShowLog
End Sub
Private Sub Command1_Click()
 If rscon.State = adStateOpen Then rscon.Close
 Unload Me
End Sub
Private Sub Command2_Click()
 PopupMenu mnu, , 4800, 7400
End Sub
Private Sub ShowLog()
 List1.Clear
 tmpval = ""
 SendMessage List1.hwnd, LB_SETTABSTOPS, MaxTab, LTab(1)
 If Combo1.ListIndex = 0 Then
  rscon.Open "SELECT a.*,b.uname FROM dbsv_adm_logfile AS a JOIN dbsv_adm_user AS b ON a.luid=b.uid OR a.luid='0';", dbcon, adOpenDynamic, adLockOptimistic
 Else
  Select Case Combo1.ListIndex
   Case 1
    tmpval = "S"
   Case 2
    tmpval = "A"
   Case 3
    tmpval = "D"
  End Select
  rscon.Open "SELECT a.*,b.uname FROM dbsv_adm_logfile AS a JOIN dbsv_adm_user AS b ON a.luid=b.uid WHERE a.ltyp='" & tmpval & "';", dbcon, adOpenDynamic, adLockOptimistic
 End If
 If rscon.EOF = True Then
  List1.AddItem "No Entries"
  Frame1.Enabled = False
 Else
  Frame1.Enabled = True
  Do While rscon.EOF = False
   List1.AddItem rscon.Fields("laktion").Value & vbTab & rscon.Fields("ldatum").Value & vbTab & IIf(rscon.Fields("luid").Value = "0", "(no user)", rscon.Fields("uname").Value)
   rscon.MoveNext
  Loop
 End If
 rscon.Close
 tmpval = ""
End Sub
Private Sub SaveLog()
 If Dir$(App.Path & "\Data\events.log") <> "" Then
  If vbYes = MsgBox("File already exists" & vbCrLf & "Overwrite?", vbExclamation + vbYesNo, MsgT) Then
   Kill App.Path & "\Data\events.log"
  Else
   Exit Sub
  End If
 End If
 rscon.Open "SELECT * FROM dbsv_adm_logfile;", dbcon, adOpenDynamic, adLockOptimistic
 fnr = FreeFile
 Open App.Path & "\Data\events.log" For Output Access Write As fnr
  Write #fnr, "Typ, Action, Date, UID"
  Do While rscon.EOF = False
   Write #fnr, rscon.Fields("ltyp").Value & ", " & rscon.Fields("laktion").Value & ", " & rscon.Fields("ldatum").Value & ", " & rscon.Fields("luid").Value
   rscon.MoveNext
  Loop
 Close fnr
 rscon.Close
End Sub
