VERSION 5.00
Begin VB.Form FrmConnDB 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Connection Data"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Kein
      Height          =   3300
      Left            =   2400
      TabIndex        =   10
      Top             =   900
      Width           =   3600
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
         ItemData        =   "FrmConnDB.frx":0000
         Left            =   1400
         List            =   "FrmConnDB.frx":000A
         Style           =   2  'Dropdown-Liste
         TabIndex        =   1
         Top             =   900
         Width           =   2000
      End
      Begin VB.TextBox Text1 
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
         Left            =   1400
         TabIndex        =   0
         Text            =   "localhost"
         Top             =   240
         Width           =   2000
      End
      Begin VB.TextBox Text3 
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
         Left            =   1400
         TabIndex        =   3
         Top             =   2100
         Width           =   2000
      End
      Begin VB.TextBox Text4 
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
         IMEMode         =   3  'DISABLE
         Left            =   1400
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2700
         Width           =   2000
      End
      Begin VB.TextBox Text2 
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
         Left            =   1400
         TabIndex        =   2
         Top             =   1500
         Width           =   2000
      End
      Begin VB.Label Label6 
         Caption         =   "Typ:"
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
         TabIndex        =   15
         Top             =   950
         Width           =   500
      End
      Begin VB.Label Label2 
         Caption         =   "Server:"
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
         TabIndex        =   14
         Top             =   285
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "User:"
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
         Left            =   105
         TabIndex        =   13
         Top             =   2150
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "Password:"
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
         Left            =   105
         TabIndex        =   12
         Top             =   2750
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "Database:"
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
         Left            =   105
         TabIndex        =   11
         Top             =   1550
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdTest 
      Height          =   660
      Left            =   6200
      Picture         =   "FrmConnDB.frx":001C
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "Test Connection"
      Top             =   3000
      Width           =   1300
   End
   Begin VB.CommandButton CmdClose 
      Height          =   660
      Left            =   3700
      Picture         =   "FrmConnDB.frx":0326
      Style           =   1  'Grafisch
      TabIndex        =   7
      ToolTipText     =   "Verlassen"
      Top             =   4500
      Width           =   1320
   End
   Begin VB.CommandButton CmdFwd 
      Height          =   660
      Left            =   5500
      Picture         =   "FrmConnDB.frx":0768
      Style           =   1  'Grafisch
      TabIndex        =   6
      ToolTipText     =   "Speichern"
      Top             =   4500
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'Kein
      Height          =   3900
      Left            =   200
      Picture         =   "FrmConnDB.frx":0BAA
      ScaleHeight     =   3900
      ScaleWidth      =   1995
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   240
      Width           =   2000
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
      Height          =   495
      Left            =   2500
      TabIndex        =   8
      Top             =   300
      Width           =   5500
   End
End
Attribute VB_Name = "FrmConnDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Sub Form_Load()
 Label1.Caption = "Enter Connection Data for the Database Server:" & vbCrLf & "(user must have privileges to that database)"
End Sub
Private Sub CmdTest_Click()
 If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Combo1.ListIndex = -1 Then
  MsgBox "Geben Sie die Verbindungsdaten an", , MsgT
  Exit Sub
 End If
 If ConnectDB = False Then
  If dbcon.State = adStateOpen Then dbcon.Close
  Text4.Text = ""
 Else
  If TableExist("dbsv_adm_logfile") = False Then
   dbcon.Close
   MsgBox "Verbindungsversuch war erfolgreich" & vbCrLf & "Die Tabellen für DBSV wurden aber nicht gefunden" & vbCrLf & "Prüfen Sie folgendes:" & vbCrLf & "* Angabe richtiger Server/Datenbank", vbCritical, MsgT
   Exit Sub
  End If
  MsgBox "Verbindungsversuch erfolgreich", vbExclamation, MsgT
  Frame1.Enabled = False
  CmdFwd.Visible = True
  CmdTest.Visible = False
 End If
End Sub
Private Sub CmdFwd_Click()
 tmpval = Write_INI("Database", "Server", Text1.Text)
 tmpval = Write_INI("Database", "Typ", Combo1.List(Combo1.ListIndex))
 tmpval = Write_INI("Database", "DB", Text2.Text)
 tmpval = Write_INI("Database", "User", Text3.Text)
 tmpval = Write_INI("Database", "Pass", Text4.Text)
 tmpval = ""
 Log_Entry "S", "Verbindung DBSV -> Datenbank erstellt", 0
 dbcon.Close
 MsgBox "Verbindungsdaten eingerichtet" & vbCrLf & "Sie können das TestModul nun verwenden", vbExclamation, MsgT
 Unload Me
End Sub
Private Sub CmdClose_Click()
 Unload Me
End Sub
Private Function ConnectDB() As Boolean
 On Error GoTo errorcon
 ConnectDB = False
 
 If Combo1.ListIndex = 0 Then
  dbcon.Open "Provider=SQL Server Native Client 10.0;Data Source='" & Text1.Text & "';Initial Catalog='" & Text2.Text & "';User ID='" & Text3.Text & "';Password='" & Text4.Text & "';"
 Else
  dbcon.Open "DRIVER={MySQL ODBC 5.1 Driver};Server=" & Text1.Text & ";UID=" & Text3.Text & ";PWD=" & Text4.Text & ";Database=" & Text2.Text
 End If
 ConnectDB = True
 Exit Function
 
errorcon:
 ConnectDB = False
 MsgBox "Could not connect to database" & vbCrLf & Err.Description, vbCritical, MsgT
End Function
Private Function TableExist(vTable As String) As Boolean
 TableExist = False
 Set rscon = dbcon.OpenSchema(adSchemaTables)
 Do While rscon.EOF = False
  If rscon!TABLE_NAME = vTable Then
   TableExist = True
   Exit Do
  End If
  rscon.MoveNext
 Loop
 rscon.Close
 Set rscon = Nothing
End Function
Private Function Write_INI(iSection As String, iKeyName As String, iValue As String)
 Dim ret As Long
 
 ret = WritePrivateProfileString(iSection, iKeyName, iValue, App.Path & "\Data\verwalt.cfg")
 Write_INI = ret
End Function
