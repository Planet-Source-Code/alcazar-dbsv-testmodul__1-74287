VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMainS 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "DBSV (Schülermodul/Tests)"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7605
   Icon            =   "FrmMainS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   7605
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.ListView ListView1 
      Height          =   1500
      Left            =   300
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   7000
      _ExtentX        =   12356
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
      Picture         =   "FrmMainS.frx":030A
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMainS.frx":34E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMainS.frx":3803
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMainS.frx":3C55
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMainS.frx":3F6F
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
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10584
            MinWidth        =   10584
            Key             =   "uID"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
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
      Begin VB.Menu mnupass 
         Caption         =   "Passwort setzen"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Beenden"
      End
   End
   Begin VB.Menu mnumod 
      Caption         =   "Module"
      Begin VB.Menu mnutest 
         Caption         =   "Tests"
      End
   End
   Begin VB.Menu mnuhlp 
      Caption         =   "Hilfe"
      Begin VB.Menu mnupinf 
         Caption         =   "Programminfo"
      End
   End
End
Attribute VB_Name = "FrmMainS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private xp As Single, yp As Single
Private Sub Form_Load()
 StatusBar1.Panels("uID").Text = "Benutzer:  " & UsrInf.Name & " ( " & tmpval & " )"
 tmpval = ""
 SetMainMnu
End Sub
Private Sub Form_Unload(Cancel As Integer)
 If vbYes = MsgBox("Schuelermodul verlassen?", vbExclamation + vbYesNo, MsgT) Then
  Log_Entry "A", "Abmeldung SchuelerModul/Tests: " & UsrInf.Name, UsrInf.ID
  dbcon.Close
 Else
  Cancel = 1
 End If
End Sub
Private Sub mnupass_Click()
 FrmPass.Show vbModal
End Sub
Private Sub mnuexit_Click()
 Unload Me
End Sub
Private Sub mnutest_Click()
 FrmTestS.Show vbModal
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
  Case "Beenden"
   mnuexit_Click
  Case "Tests"
   mnutest_Click
  Case "Programminfo"
   mnupinf_Click
 End Select
End Sub
Private Sub SetMainMnu()
 With ListView1.ListItems
  .Clear
  Set listX = .Add(, , "Passwort setzen", 1)
  Set listX = .Add(, , "Tests", 3)
  Set listX = .Add(, , "Programminfo", 4)
  Set listX = .Add(, , "Beenden", 2)
 End With
End Sub
