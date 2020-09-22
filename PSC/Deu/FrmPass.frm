VERSION 5.00
Begin VB.Form FrmPass 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Passwort ändern"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      Height          =   2700
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   6200
      Begin VB.CommandButton Command2 
         Height          =   600
         Left            =   4700
         Picture         =   "FrmPass.frx":0000
         Style           =   1  'Grafisch
         TabIndex        =   4
         Top             =   1100
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Height          =   600
         Left            =   4700
         Picture         =   "FrmPass.frx":030A
         Style           =   1  'Grafisch
         TabIndex        =   3
         Top             =   300
         Visible         =   0   'False
         Width           =   1200
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
         IMEMode         =   3  'DISABLE
         Left            =   1900
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1900
         Visible         =   0   'False
         Width           =   2200
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
         IMEMode         =   3  'DISABLE
         Left            =   1900
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1200
         Visible         =   0   'False
         Width           =   2200
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
         IMEMode         =   3  'DISABLE
         Left            =   1900
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   500
         Width           =   2200
      End
      Begin VB.Label Label3 
         Caption         =   "Passwort wiederholen:"
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
         Left            =   150
         TabIndex        =   8
         Top             =   1850
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Neues Passwort:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   1250
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Altes Passwort:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   550
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If Text1.Text <> "" Then
   rscon.Open "SELECT upass FROM dbsv_adm_user WHERE uid='" & UsrInf.ID & "' AND upass='" & Text1.Text & "';", dbcon, adOpenDynamic, adLockOptimistic
   If rscon.EOF = False Then
    rscon.Close
    Label1.Visible = False
    Label2.Visible = True
    Label3.Visible = True
    Text1.Visible = False
    Text2.Visible = True
    Text3.Visible = True
    Text2.SetFocus
    Command1.Visible = True
    Exit Sub
   End If
   rscon.Close
  End If
  MsgBox "Ungültiges Passwort", vbExclamation, MsgT
  Text1.Text = ""
  Text1.SetFocus
 End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Text3.SetFocus
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Command1_Click
End Sub
Private Sub Command1_Click()
 If LCase$(Text1.Text) = LCase$(Text2.Text) Then
  MsgBox "Das neue Passwort muß sich vom alten unterscheiden", vbExclamation, MsgT
  SetItem
  Exit Sub
 End If
 If IsValPw(UsrInf.Name, Text2.Text, Text3.Text) = True Then
  dbcon.Execute "UPDATE dbsv_adm_user SET upass='" & Text3.Text & "' WHERE uid='" & UsrInf.ID & "';"
  Log_Entry "D", "Passwortänderung: " & UsrInf.Name, UsrInf.ID
  Command2_Click
 Else
  SetItem
 End If
End Sub
Private Sub Command2_Click()
 Unload Me
End Sub
Private Sub SetItem()
 Text2.Text = ""
 Text3.Text = ""
 Text2.SetFocus
End Sub
