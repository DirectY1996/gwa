VERSION 5.00
Begin VB.Form frmBot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Bot"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4980
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Selecionar Novo Status"
      Height          =   2115
      Left            =   2580
      TabIndex        =   4
      Top             =   780
      Width           =   2235
      Begin VB.OptionButton Option1 
         Caption         =   "Reiniciando"
         Height          =   375
         Index           =   2
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1380
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "On"
         Height          =   375
         Index           =   1
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   420
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Off"
         Height          =   375
         Index           =   0
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   900
         Width           =   1155
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   360
      Index           =   1
      Left            =   2700
      TabIndex        =   3
      Top             =   3720
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   360
      Index           =   0
      Left            =   3780
      TabIndex        =   2
      Top             =   3720
      Width           =   990
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numeros Associados a Este Bot:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   2310
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public BotSeq As Long

Dim StatusMudou As Boolean

Private Sub Command1_Click(Index As Integer)
Dim Z As Long
If Index = 0 And StatusMudou Then
    If Option1(0).Value Then
    Z = 0
    ElseIf Option1(1).Value Then
    Z = 1
    ElseIf Option1(2).Value Then
    Z = 2
    Else
    Unload Me
    Exit Sub
    End If
    DBFire.Execute "UPDATE TB_BOT SET BOT_STATUS = " & Z & " WHERE BOT_SEQ = " & BotSeq
End If
Unload Me
End Sub

Private Sub Form_Load()
Dim F As New hDirectFire, Z As Long
List1.Clear
If F.OpenRecordset("SELECT BOT_SEQ,BOT_STATUS FROM TB_BOT WHERE BOT_SEQ = " & BotSeq & " ORDER BY BOT_SEQ") Then
    Z = F("BOT_STATUS")
    If Z >= 0 And Z <= 2 Then Option1(Z).Value = True
End If
F.CloseTable
If F.OpenRecordset("SELECT WAS_SEQ,WAS_VAL FROM TB_WAS WHERE WAS_BOT_SEQ = " & BotSeq & " ORDER BY WAS_SEQ") Then
    Do Until F.EOF
        List1.AddItem F("WAS_SEQ") & ":" & F("WAS_VAL")
        F.MoveNext
    Loop
End If
End Sub

Private Sub Option1_Click(Index As Integer)
StatusMudou = True
End Sub
