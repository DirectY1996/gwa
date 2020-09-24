VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public SesSeq As Long
Public BotSeq As Long
Public WasCli As String 'Contato
Public WasBot As String 'Contato

Event SendMessage(Text As String)

Sub Create()
SesSeq = DBFire.GetGenerator("GE_SES_SEQ")
DBFire.Execute "INSERT INTO TB_SES (SES_SEQ,SES_BOT_SEQ,SES_ATD_SEQ,SES_CLI,SES_WAS) VALUES (" & SesSeq & "," & BotSeq & ",NULL," & WasCli & "," & WasBot & ")"
End Sub

Sub Destroy()
DBFire.Execute "DELETE FROM TB_SES WHERE SES_SEQ = " & SesSeq
End Sub

Function Tick() As Boolean 'Return True To Live Another Day, Return False To Die
Tick = False
End Function

Sub PutMessage(Text As String)
RaiseEvent SendMessage("Você Disse:" & Text)
End Sub
