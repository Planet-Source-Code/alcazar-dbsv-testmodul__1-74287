VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "DBSV (Hauptmenue)"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9405
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   9405
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.ListView ListView1 
      Height          =   3000
      Left            =   300
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   8800
      _ExtentX        =   15531
      _ExtentY        =   5292
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
      Picture         =   "FrmMain.frx":030A
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
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":34E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3803
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":44DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":492F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4C49
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4F63
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":527D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Unten ausrichten
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   3525
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14112
            MinWidth        =   14112
            Key             =   "uID"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "02.07.2012"
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
      Begin VB.Menu mnupass 
         Caption         =   "Passwort setzen"
      End
      Begin VB.Menu mnulf 
         Caption         =   "LogDatei"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Beenden"
      End
   End
   Begin VB.Menu mnusdata 
      Caption         =   "Schuldaten"
      Begin VB.Menu mnuf 
         Caption         =   "Fächer"
      End
      Begin VB.Menu mnuk 
         Caption         =   "Klassen"
      End
   End
   Begin VB.Menu mnumod 
      Caption         =   "Module"
      Begin VB.Menu mnutestedit 
         Caption         =   "TestEditor"
      End
   End
   Begin VB.Menu mnuhlp 
      Caption         =   "Hilfe"
      Begin VB.Menu mnupinf 
         Caption         =   "Programminfo"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private xp As Single, yp As Single
Private Sub Form_Load()
 StatusBar1.Panels("uID").Text = "Benutzer:  " & UsrInf.Name & " ( " & UsrInf.RName & " )"
 SetMnu
End Sub
Private Sub Form_Unload(Cancel As Integer)
 If vbYes = MsgBox("Objektmanager schließen und DBSV verlassen?", vbExclamation + vbYesNo, MsgT) Then
  Log_Entry "A", "Abmeldung DBSV: " & UsrInf.Name, UsrInf.ID
  dbcon.Close
 Else
  Cancel = 1
 End If
End Sub
Private Sub mnupass_Click()
 FrmPass.Show vbModal
End Sub
Private Sub mnulf_Click()
 FrmLogFile.Show vbModal
End Sub
Private Sub mnuexit_Click()
 Unload Me
End Sub
Private Sub mnuf_Click()
 FrmFach.Show vbModal
End Sub
Private Sub mnuk_Click()
 FrmCls.Show vbModal
End Sub
Private Sub mnutestedit_Click()
 FrmTest.Show vbModal
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
  Case "Passwort setzen"
   mnupass_Click
  Case "LogDatei"
   mnulf_Click
  Case "Beenden"
   mnuexit_Click
  Case "Fächer"
   mnuf_Click
  Case "Klassen"
   mnuk_Click
  Case "TestEditor"
   mnutestedit_Click
  Case "Programminfo"
   mnupinf_Click
 End Select
End Sub
Private Sub SetMnu()
 With UsrInf
  If Mid$(.ACL, 1, 1) = "0" Then mnulf.Visible = False
  If Val(Mid$(.ACL, 2, 2)) = "0" Then
   mnusdata.Visible = False
  Else
   If Mid$(.ACL, 2, 1) = "0" Then mnuf.Visible = False
   If Mid$(.ACL, 3, 1) = "0" Then mnuk.Visible = False
  End If
  If Mid$(.ACL, 4, 1) = "0" Then mnumod.Visible = False
 End With
 SetMainMnu
End Sub
Private Sub SetMainMnu()
 With ListView1.ListItems
  .Clear
  Set listX = .Add(, , "Passwort setzen", 1)
  If Mid$(UsrInf.ACL, 1, 1) = "1" Then Set listX = .Add(, , "LogDatei", 2)
  If Mid$(UsrInf.ACL, 2, 1) = "1" Then Set listX = .Add(, , "Fächer", 4)
  If Mid$(UsrInf.ACL, 3, 1) = "1" Then Set listX = .Add(, , "Klassen", 5)
  If Mid$(UsrInf.ACL, 4, 1) = "1" Then Set listX = .Add(, , "TestEditor", 6)
  Set listX = .Add(, , "Programminfo", 7)
  Set listX = .Add(, , "Beenden", 3)
 End With
End Sub
