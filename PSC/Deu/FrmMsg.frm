VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmMsg 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Fehler:"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   1000
      Top             =   2000
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'Kein
      Height          =   1500
      Left            =   240
      Picture         =   "FrmMsg.frx":0000
      ScaleHeight     =   1500
      ScaleWidth      =   615
      TabIndex        =   1
      Top             =   200
      Width           =   615
   End
   Begin VB.Label Label1 
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
      Height          =   1500
      Left            =   1200
      TabIndex        =   2
      Top             =   255
      Width           =   5700
   End
   Begin MSForms.CommandButton Command1 
      Height          =   510
      Left            =   3000
      TabIndex        =   0
      Top             =   1900
      Width           =   1300
      Caption         =   "OK"
      Size            =   "2302;900"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "FrmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
 Timer1.Enabled = False
 If Command1.Tag = 1 Then dbcon.Close
 Unload Me
End Sub
Private Sub Timer1_Timer()
 Command1_Click
End Sub
Function LoadFrm(vClose As Integer)
 Command1.Tag = vClose
 Label1.Caption = "Das System konnte Ihre Anmeldung nicht bestaetigen!" & vbCrLf & "Prüfen Sie folgendes:" & vbCrLf & " * Richtige Eingabe von Benutzername und Passwort" & vbCrLf & " * Ob Ihre Benutzer-ID aktiviert ist" & vbCrLf & " * Ob Ihre Rolle Berechtigungen für den Programmteil besitzt"
 Me.Show
End Function
