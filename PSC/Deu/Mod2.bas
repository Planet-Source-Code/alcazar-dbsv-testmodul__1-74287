Attribute VB_Name = "Mod2"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Function OpenDB() As Boolean
 Dim mServer As String, mTyp As String, mDB As String, mUser As String, mPass As String
 
 On Error GoTo errorcon
 OpenDB = False
 
 mServer = Read_INI("Database", "Server")
 mTyp = Read_INI("Database", "Typ")
 mDB = Read_INI("Database", "DB")
 mUser = Read_INI("Database", "User")
 mPass = Read_INI("Database", "Pass")
 If mTyp = "MSSQL" Then
  dbcon.Open "Provider=SQL Server Native Client 10.0;Data Source='" & mServer & "';Initial Catalog='" & mDB & "';User ID='" & mUser & "';Password='" & mPass & "';"
 Else
  dbcon.Open "DRIVER={MySQL ODBC 5.1 Driver};Server=" & mServer & ";UID=" & mUser & ";PWD=" & mPass & ";Database=" & mDB
 End If
 OpenDB = True
 mPass = ""
 Exit Function
 
errorcon:
 OpenDB = False
 MsgBox "Could not connect to database" & vbCrLf & Err.Description, vbCritical, MsgT
End Function
Private Function Read_INI(iSection As String, iKeyName As String)
 Dim ret As String, ret2 As Long

 ret = String(255, 0)
 ret2 = GetPrivateProfileString(iSection, iKeyName, "", ret, Len(ret), App.Path & "\Data\verwalt.cfg")
 If ret2 <> 0 Then ret = Left$(ret, ret2) Else ret = ""
 Read_INI = ret
End Function
Function Log_Entry(Typ As String, LAkt As String, LVon As Integer)
 Dim tmpdate As String

 tmpdate = CStr(Format(Date, "yyyy-mm-dd"))
 dbcon.Execute "INSERT INTO dbsv_adm_logfile (ltyp,laktion,ldatum,luid) VALUES('" & Typ & "','" & LAkt & "','" & tmpdate & "','" & LVon & "');"
End Function
Function IsValPw(Usr As String, Pw As String, Pw2 As String) As Boolean
 IsValPw = False
 If Pw = "" Then
  MsgBox "Bitte geben Sie ein Passwort ein", vbExclamation, MsgT
  Exit Function
 End If
 If Len(Pw) < 6 Then
  MsgBox "Das Passwort muß aus mindestens 6 Zeichen bestehen", vbExclamation, MsgT
  Exit Function
 End If
 If LCase$(Pw) = LCase$(Usr) Then
  MsgBox "Der Benutzername kann nicht das Passwort sein", vbExclamation, MsgT
  Exit Function
 End If
 If Pw = Pw2 Then
  IsValPw = True
 Else
  MsgBox "Das Passwort wurde nicht korrekt bestätigt", vbExclamation, MsgT
 End If
End Function
Function ValID(vID As String, vTable As String, vField As String) As Boolean
 ValID = False
 If vID = "" Then
  MsgBox "Objektnamen eingeben", vbExclamation, MsgT
  Exit Function
 End If
 If LCase$(vID) = "intendant" Or LCase$(vID) = "system" Then
  MsgBox "Objektname reserviert, verwenden Sie einen anderen Namen", vbExclamation, MsgT
  Exit Function
 End If
 rscon.Open "SELECT " & vField & " FROM " & vTable & " WHERE " & vField & "='" & vID & "';", dbcon, adOpenDynamic, adLockOptimistic
 If rscon.EOF = True Then
  rscon.Close
  ValID = True
 Else
  rscon.Close
  MsgBox "Objektname bereits vorhanden", vbExclamation, MsgT
 End If
End Function
Private Sub Main()
 If App.PrevInstance = False Then
  If Dir$(App.Path & "\Data\verwalt.cfg") = "" Then
   MsgBox "Konfigurationsdatei nicht gefunden" & vbCrLf & "Verbindung zur Datenbank nicht möglich" & vbCrLf & "Geben Sie die Verbindungsdaten an", vbExclamation, MsgT
   FrmConnDB.Show
   Exit Sub
  End If
  If OpenDB = True Then
   eröffnung.Show
  Else
   If dbcon.State = adStateOpen Then dbcon.Close
  End If
 End If
End Sub
