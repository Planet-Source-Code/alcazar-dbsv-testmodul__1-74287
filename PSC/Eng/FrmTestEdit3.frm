VERSION 5.00
Begin VB.Form FrmTestEdit3 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "TestEditor - Sort Questions"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton CmdCancel 
      Height          =   650
      Left            =   5300
      Picture         =   "FrmTestEdit3.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   7300
      Width           =   1300
   End
   Begin VB.CommandButton CmdSave 
      Height          =   650
      Left            =   3300
      Picture         =   "FrmTestEdit3.frx":030A
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   7300
      Width           =   1300
   End
   Begin VB.ListBox List2 
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
      Left            =   5200
      TabIndex        =   1
      Top             =   900
      Width           =   4500
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
      Left            =   300
      TabIndex        =   0
      Top             =   900
      Width           =   4500
   End
   Begin VB.Label Label2 
      Caption         =   "sorted questions:"
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
      Left            =   5300
      TabIndex        =   5
      Top             =   400
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "available questions:"
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
      Top             =   400
      Width           =   2100
   End
End
Attribute VB_Name = "FrmTestEdit3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub List1_DblClick()
 If List1.ListIndex <> -1 Then
  List2.AddItem List1.List(List1.ListIndex)
  List2.ItemData(List2.NewIndex) = List1.ItemData(List1.ListIndex)
  List1.RemoveItem List1.ListIndex
 End If
End Sub
Private Sub List2_DblClick()
 If List2.ListIndex <> -1 Then
  List1.AddItem List2.List(List2.ListIndex)
  List1.ItemData(List1.NewIndex) = List2.ItemData(List2.ListIndex)
  List2.RemoveItem List2.ListIndex
 End If
End Sub
Private Sub CmdSave_Click()
 If List1.ListCount <> 0 Then
  MsgBox "The left list still contains questions to sort.", vbExclamation, MsgT
 Else
  For z = 0 To List2.ListCount - 1
   dbcon.Execute "UPDATE dbsv_test_cat SET tcorder='" & z & "' WHERE tcid='" & List2.ItemData(z) & "';"
  Next
  Unload Me
 End If
End Sub
Private Sub CmdCancel_Click()
 Unload Me
End Sub
