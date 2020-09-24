VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecionar Empresa"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3210
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
   ScaleHeight     =   6450
   ScaleWidth      =   3210
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Criar Bot"
      Height          =   915
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2955
      Begin VB.CommandButton Command3 
         Caption         =   "Criar Bot"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1860
         TabIndex        =   7
         Top             =   480
         Width           =   990
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bots Offline:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Tag             =   "Bots Offline: "
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecionar Empresa"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2955
      Begin VB.ListBox List1 
         Height          =   4350
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2715
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         Height          =   360
         Left            =   2400
         TabIndex        =   2
         Top             =   4680
         Width           =   450
      End
      Begin VB.CommandButton Command2 
         Caption         =   "-"
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   4680
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
Command3.Enabled = True
End Sub

Private Sub Command1_Click()
DBFire.Execute "INSERT INTO TB_EMP (EMP_SEQ,EMP_NAME) VALUES (GEN_ID(GE_EMP_SEQ,1),'Nova Empresa')"
RefreshLists
End Sub

Private Sub Command2_Click()
Dim C() As String
C = Split(List1.List(List1.ListIndex), ":")
DBFire.Execute "DELETE FROM TB_EMP WHERE EMP_SEQ = " & C(0)
RefreshLists
End Sub

Private Sub Command3_Click()
Dim S As String
S = Combo1.List(Combo1.ListIndex)
frmIntel.BotSeq = CLng(Left$(S, Len(S) - 4)) ':Off'
On Error Resume Next
frmIntel.Show
On Error GoTo 0
If frmIntel.BotSeq Then Unload Me
End Sub

Private Sub Form_Load()
RefreshLists
End Sub

Private Sub RefreshLists()
Dim F As New hDirectFire
List1.Clear
If F.OpenRecordset("SELECT EMP_SEQ,EMP_NAME FROM TB_EMP ORDER BY EMP_SEQ") Then
    Do Until F.EOF
    List1.AddItem F("EMP_SEQ") & ": " & F("EMP_NAME")
    F.MoveNext
    Loop
End If
F.CloseTable
Combo1.Clear
If F.OpenRecordset("SELECT BOT_SEQ FROM TB_BOT WHERE BOT_STATUS = 0 ORDER BY BOT_SEQ") Then
    Do Until F.EOF
    Combo1.AddItem F("BOT_SEQ") & ":Off"
    F.MoveNext
    Loop
End If
Label1.Caption = Label1.Tag & Combo1.ListCount & " Bots"
Command2.Enabled = False
End Sub

Private Sub List1_Click()
Command2.Enabled = True
End Sub

Private Sub List1_DblClick()
Dim C() As String
C = Split(List1.List(List1.ListIndex), ":")
frmEmp.Show
frmEmp.Present (CLng(C(0)))
Unload Me
End Sub
