VERSION 5.00
Begin VB.Form FrmClsDet 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Klassen"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command1 
      Height          =   660
      Left            =   1700
      Picture         =   "FrmClsDet.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   1900
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Height          =   660
      Left            =   3500
      Picture         =   "FrmClsDet.frx":030A
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   1900
      Width           =   1320
   End
   Begin VB.TextBox TxtName 
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
      Left            =   2200
      MaxLength       =   25
      TabIndex        =   0
      Top             =   350
      Width           =   4000
   End
   Begin VB.TextBox TxtDescr 
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
      Left            =   2200
      MaxLength       =   255
      TabIndex        =   1
      Top             =   1050
      Width           =   4000
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
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
      Top             =   400
      Width           =   700
   End
   Begin VB.Label Label2 
      Caption         =   "Beschreibung:"
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
      Top             =   1100
      Width           =   1400
   End
End
Attribute VB_Name = "FrmClsDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
 tmpval = ""
 If ValID(TxtName.Text, "dbsv_main_class", "cname") = False Then Exit Sub
 dbcon.Execute "INSERT INTO dbsv_main_class (cname,descr) VALUES('" & TxtName.Text & "','" & IIf(TxtDescr.Text = "", "", TxtDescr.Text) & "');"
 rscon.Open "SELECT last_insert_id();", dbcon, adOpenDynamic, adLockOptimistic
 tmpval = rscon.Fields("last_insert_id()").Value
 rscon.Close
 Set nodX = FrmCls.TreeView1.Nodes.Add("root", tvwChild, tmpval & "x", TxtName.Text, 2)
 FrmCls.TreeView1.Refresh
 Log_Entry "D", "Erstellung KLasse: " & TxtName.Text, UsrInf.ID
 Command2_Click
End Sub
Private Sub Command2_Click()
 tmpval = ""
 Unload Me
End Sub
