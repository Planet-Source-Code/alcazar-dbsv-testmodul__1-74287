VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form eröffnung 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Willkommen"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   ControlBox      =   0   'False
   Icon            =   "eröffnung.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      Height          =   5000
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   8100
      Begin VB.Timer Timer1 
         Interval        =   30
         Left            =   7500
         Top             =   300
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'Kein
         Height          =   1550
         Left            =   300
         Picture         =   "eröffnung.frx":030A
         ScaleHeight     =   1545
         ScaleWidth      =   600
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'Kein
         Height          =   1550
         Left            =   300
         Picture         =   "eröffnung.frx":0AAC
         ScaleHeight     =   1545
         ScaleWidth      =   600
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   600
      End
      Begin MSForms.CommandButton Command1 
         Height          =   700
         Index           =   0
         Left            =   500
         TabIndex        =   0
         Top             =   3800
         Width           =   2100
         Caption         =   "  DBSV"
         PicturePosition =   327683
         Size            =   "3704;1244"
         Picture         =   "eröffnung.frx":162A
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton Command1 
         Height          =   700
         Index           =   1
         Left            =   2950
         TabIndex        =   1
         Top             =   3800
         Width           =   2100
         Caption         =   "  Verwaltung"
         PicturePosition =   327683
         Size            =   "3704;1235"
         Picture         =   "eröffnung.frx":1944
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton Command2 
         Height          =   700
         Left            =   5400
         TabIndex        =   2
         Top             =   3800
         Width           =   2100
         Caption         =   "Schliessen"
         PicturePosition =   327683
         Size            =   "3704;1235"
         Picture         =   "eröffnung.frx":1C5E
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label6 
         Caption         =   $"eröffnung.frx":1F78
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   1300
         TabIndex        =   6
         Top             =   2100
         Width           =   5800
      End
      Begin VB.Label Label5 
         Caption         =   "Willkommen zum DBSV-Verwaltungssystem !"
         BeginProperty Font 
            Name            =   "Arioso"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   900
         TabIndex        =   5
         Top             =   600
         Width           =   6400
      End
      Begin VB.Label Label2 
         Caption         =   "Das System heißt Sie willkommen."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2200
         TabIndex        =   4
         Top             =   1300
         Width           =   3700
      End
   End
End
Attribute VB_Name = "eröffnung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
 eröffnung.Width = 200
End Sub
Private Sub Command1_Click(Index As Integer)
 Unload Me
 FrmKeyM.LoadFrm Index
End Sub
Private Sub Command2_Click()
 dbcon.Close
 Unload Me
End Sub
Private Sub Timer1_Timer()
 If eröffnung.Width >= 8700 Then
  Timer1.Enabled = False
  eröffnung.Width = 8700
 Else
  eröffnung.Left = eröffnung.Left - 250
  eröffnung.Width = eröffnung.Width + 500
 End If
End Sub
