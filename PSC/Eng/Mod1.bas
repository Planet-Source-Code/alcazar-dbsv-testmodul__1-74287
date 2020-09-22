Attribute VB_Name = "Mod1"
Option Explicit
Public nodX As Node, listX As ListItem, z As Integer, tmpval As String
Public dbcon As New ADODB.Connection, rscon As New ADODB.Recordset
Public Const MsgT As String = "The System:"
Private Type UsrData
 ID As Integer
 Name As String
 Role As Integer
 RName As String
 ACL As String
 Class As Integer
End Type
Private Type TestData
 TimeL As Integer
 TimeExp As Byte
 ScoreToPass As Integer
End Type
Public UsrInf As UsrData, TestInf As TestData
