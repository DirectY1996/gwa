VERSION 5.00
Begin VB.Form frmIntel 
   Caption         =   "Bot #"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   420
      Top             =   660
   End
End
Attribute VB_Name = "frmIntel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public BotSeq As Long

Dim S() As clsSes
Dim MaxS As Long

Dim WithEvents C As clsSes
Attribute C.VB_VarHelpID = -1

Dim MsgFilter As String

Private Sub Log(Text As String)
List1.AddItem Text
List1.ListIndex = List1.ListCount - 1
End Sub

Private Sub Form_Load()
Caption = "Bot #" & BotSeq
Dim F As New hDirectFire, Z As Long
MaxS = -1
DBFire.Execute "UPDATE TB_BOT SET BOT_STATUS = 1 WHERE BOT_SEQ = " & BotSeq
If F.OpenRecordset("SELECT WAS_SEQ,WAS_VAL FROM TB_WAS WHERE WAS_BOT_SEQ = " & BotSeq & " ORDER BY WAS_SEQ") Then
    If F.EOF Then
        MsgBox "Este Bot Não Tem Nenhum Numero Associado a Ele!"
        BotSeq = 0
        On Error Resume Next
        Unload Me
        Exit Sub
    End If
    Do Until F.EOF
        MsgFilter = MsgFilter & "MSG_BOT = '" & F("WAS_VAL") & "' OR "
        F.MoveNext
    Loop
    MsgFilter = "MSG_STATUS = 0 AND MSG_DIR = 0 AND MSG_TYPE = 1 AND (" & Left$(MsgFilter, Len(MsgFilter) - 4) & ")"
Else
    MsgBox F.LastErrDescription
    On Error Resume Next
    Unload Me
    Exit Sub
End If
End Sub

Private Sub Form_Resize()
List1.Height = ScaleHeight
List1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Z As Long
For Z = 0 To MaxS
If Not S(Z) Is Nothing Then S(Z).Destroy
Next
DBFire.Execute "UPDATE TB_BOT SET BOT_STATUS = 0 WHERE BOT_SEQ = " & BotSeq
End Sub

Private Sub Timer1_Timer()
Dim F As New hDirectFire, Z As Long
If F.OpenRecordset("SELECT MSG_SEQ,MSG_CLI,MSG_BOT,MSG_DATA FROM TB_MSG WHERE " & MsgFilter & " ORDER BY MSG_SEQ") Then
    DBFire.Execute "UPDATE TB_MSG SET MSG_STATUS = 1 WHERE " & MsgFilter
    Do Until F.EOF
    TratarMsg StrConv(F("MSG_DATA"), vbUnicode), F("MSG_CLI"), F("MSG_BOT") 'Novas Mensagens
    F.MoveNext
    Loop
Else
Timer1.Enabled = False
MsgBox F.LastErrDescription
End If
For Z = 0 To MaxS
    If Not S(Z) Is Nothing Then
        Set C = S(Z)
        If Not C.Tick Then
        Log "Sessão " & C.SesSeq & " Terminada"
        C.Destroy
        Set C = Nothing
        Set S(Z) = Nothing
        Else
        Set C = Nothing
        End If
    End If
Next
End Sub

Private Sub TratarMsg(Text As String, WasCli As String, WasBot As String)
Dim Z As Long
For Z = 0 To MaxS
    If Not S(Z) Is Nothing Then
        If S(Z).WasCli = WasCli And S(Z).WasBot = WasBot Then
        Set C = S(Z)
        GoTo OK
        End If
    End If
Next
For Z = 0 To MaxS 'Criar um novo Bot
If S(Z) Is Nothing Then Exit For
Next
If Z > MaxS Then
MaxS = Z
ReDim Preserve S(Z)
End If
Set S(Z) = New clsSes
Set C = S(Z)
C.BotSeq = BotSeq
C.WasBot = WasBot
C.WasCli = WasCli
C.Create
Log "Sessão " & C.SesSeq & " Criada"
OK:
Log Text
C.PutMessage Text
Set C = Nothing
End Sub

Private Sub C_SendMessage(Text As String)
DBFire.Execute "INSERT INTO TB_MSG (MSG_SEQ,MSG_CLI,MSG_BOT,MSG_DATA,MSG_TYPE,MSG_STATUS,MSG_DIR) VALUES (GEN_ID(GE_MSG_SEQ,1),'" & C.WasCli & "','" & C.WasBot & "','" & Text & "',1,0,1)"
End Sub
