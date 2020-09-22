VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMainA 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "DBSV (Verwaltung)"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9405
   Icon            =   "FrmMainA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.ListView ListView1 
      Height          =   1500
      Left            =   300
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   8800
      _ExtentX        =   15531
      _ExtentY        =   2646
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HotTracking     =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
      Picture         =   "FrmMainA.frx":0442
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   100
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMainA.frx":3621
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMainA.frx":3A73
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMainA.frx":3D8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMainA.frx":40A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMainA.frx":4D81
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Unten ausrichten
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   2025
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11466
            MinWidth        =   11466
            Key             =   "uID"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "27.06.2012"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "Datei"
      Begin VB.Menu mnuexit 
         Caption         =   "Beenden"
      End
   End
   Begin VB.Menu mnusec 
      Caption         =   "Sicherheit"
      Begin VB.Menu mnusecuser 
         Caption         =   "Benutzer"
      End
      Begin VB.Menu mnusecrole 
         Caption         =   "Rollen"
      End
      Begin VB.Menu mnulogfile 
         Caption         =   "LogDatei"
      End
   End
   Begin VB.Menu mnuhlp 
      Caption         =   "Hilfe"
      Begin VB.Menu mnupinf 
         Caption         =   "Programminfo"
      End
   End
End
Attribute VB_Name = "FrmMainA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private xp As Single, yp As Single
Private Sub Form_Load()
 StatusBar1.Panels("uID").Text = "Benutzer:  " & UsrInf.Name & " ( " & UsrInf.RName & " )"
 SetMainMnu
End Sub
Private Sub Form_Unload(Cancel As Integer)
 If vbYes = MsgBox("Administration schließen und DBSV verlassen?", vbExclamation + vbYesNo, MsgT) Then
  Log_Entry "A", "Abmeldung Administration: " & UsrInf.Name, UsrInf.ID
  dbcon.Close
 Else
  Cancel = 1
 End If
End Sub
Private Sub mnuexit_Click()
 Unload Me
End Sub
Private Sub mnusecuser_Click()
 FrmUsrList.Show vbModal
End Sub
Private Sub mnusecrole_Click()
 FrmRoleList.Show vbModal
End Sub
Private Sub mnulogfile_Click()
 FrmLogFile.Show vbModal
End Sub
Private Sub mnupinf_Click()
 prginfo.Show vbModal
End Sub
Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 xp = x
 yp = y
End Sub
Private Sub ListView1_Click()
 If ListView1.HitTest(xp, yp) Is Nothing Then Exit Sub
 Select Case ListView1.HitTest(xp, yp).Text
  Case "Beenden"
   mnuexit_Click
  Case "Benutzer"
   mnusecuser_Click
  Case "Rollen"
   mnusecrole_Click
  Case "LogDatei"
   mnulogfile_Click
  Case "Programminfo"
   mnupinf_Click
 End Select
End Sub
Private Sub SetMainMnu()
 With ListView1.ListItems
  .Clear
  Set listX = .Add(, , "Benutzer", 2)
  Set listX = .Add(, , "Rollen", 3)
  Set listX = .Add(, , "LogDatei", 4)
  Set listX = .Add(, , "Programminfo", 5)
  Set listX = .Add(, , "Beenden", 1)
 End With
End Sub
