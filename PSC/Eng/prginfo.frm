VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form prginfo 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Programminfo"
   ClientHeight    =   5880
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5565
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4058.48
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   5225.823
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frminfo 
      Height          =   5361
      Left            =   256
      TabIndex        =   0
      Top             =   240
      Width           =   5058
      Begin VB.Frame FrmUser 
         Caption         =   "registered for:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1300
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   4575
         Begin VB.Label Label4 
            Alignment       =   2  'Zentriert
            Caption         =   "Planet Source Code"
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
            Left            =   120
            TabIndex        =   8
            Top             =   550
            Width           =   4335
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'Kein
         Height          =   600
         Left            =   500
         Picture         =   "prginfo.frx":0000
         ScaleHeight     =   600
         ScaleWidth      =   405
         TabIndex        =   1
         Top             =   500
         Width           =   400
      End
      Begin VB.Label Label7 
         Caption         =   "Ò"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3100
         TabIndex        =   12
         Top             =   3100
         Width           =   225
      End
      Begin MSForms.CommandButton Command1 
         Height          =   600
         Left            =   1800
         TabIndex        =   11
         Top             =   4400
         Width           =   1300
         Caption         =   "Schließen"
         PicturePosition =   327683
         Size            =   "2293;1058"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label12 
         Caption         =   "Ó"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   3700
         Width           =   225
      End
      Begin VB.Label Label3 
         Caption         =   "1.20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4300
         TabIndex        =   9
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label11 
         Caption         =   "Ò"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3250
         TabIndex        =   6
         Top             =   3600
         Width           =   225
      End
      Begin VB.Label Label8 
         Caption         =   "1996 - 2012  by Evil Inc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   5
         Top             =   3800
         Width           =   2700
      End
      Begin VB.Label Label6 
         Caption         =   "Created with: Visual Basic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   3240
         Width           =   2800
      End
      Begin VB.Label Label2 
         Caption         =   "- TestModul -"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Programmsystem DBSV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   2685
      End
   End
End
Attribute VB_Name = "prginfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
 Unload Me
End Sub
Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 2 Then MsgBox "Nice Try, but the Easter Egg is only available in the DBSV Project...", vbExclamation, MsgT
End Sub
