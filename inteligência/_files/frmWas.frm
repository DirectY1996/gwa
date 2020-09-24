VERSION 5.00
Begin VB.Form frmWas 
   Caption         =   "Cadastro de Watsapp"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5490
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
   ScaleHeight     =   5070
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   360
      Index           =   1
      Left            =   3180
      TabIndex        =   6
      Top             =   4620
      Width           =   1050
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   360
      Index           =   0
      Left            =   4320
      TabIndex        =   5
      Top             =   4620
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   360
      Width           =   2835
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Descelecionar"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   4620
      Width           =   1170
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seq do Bot Selecionado:"
      Height          =   195
      Left            =   2520
      TabIndex        =   7
      Tag             =   "Seq do Bot Selecionado:"
      Top             =   780
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero:"
      Height          =   195
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bots:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmWas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Empresa As Long
Public WasSeq As Long

Private Property Get GetSeq(Optional ByVal Index As Long = -1) As Long
Dim C() As String
If Index < 0 Then Index = List1.ListIndex
If Index < 0 Then Exit Property
C = Split(List1.List(Index), ":")
On Error Resume Next
GetSeq = C(0)
End Property

Private Sub RefreshList()
Dim F As New hDirectFire, Z As Long
List1.Clear
If F.OpenRecordset("SELECT WAS_SEQ,WAS_BOT_SEQ,WAS_VAL FROM TB_WAS WHERE WAS_SEQ = " & WasSeq & " ORDER BY WAS_SEQ") Then
    If Not F.EOF Then
        Z = F("WAS_BOT_SEQ")
        Text1 = F("WAS_VAL")
        If Text1 = "0" Then Text1 = ""
    End If
End If
F.CloseTable
If F.OpenRecordset("SELECT BOT_SEQ,BOT_STATUS FROM TB_BOT WHERE BOT_EMP_SEQ = " & Empresa & " ORDER BY BOT_SEQ") Then
    Do Until F.EOF
        Select Case F("BOT_STATUS")
        Case 0: List1.AddItem F("BOT_SEQ") & ":Off"
        Case 1: List1.AddItem F("BOT_SEQ") & ":On"
        Case 2: List1.AddItem F("BOT_SEQ") & ":Reiniciando..."
        Case Else
        List1.AddItem F("BOT_SEQ") & ":?(" & CStr(F("BOT_STATUS")) & ")"
        End Select
        If F("BOT_SEQ") = Z Then
        List1.ListIndex = List1.ListCount - 1
        End If
    F.MoveNext
    Loop
End If
End Sub

Private Sub Command2_Click(Index As Integer)
If Index = 0 Then 'Update Values And Move On
    DBFire.Execute "UPDATE TB_WAS SET WAS_VAL = '" & Text1 & "' WHERE WAS_SEQ = " & WasSeq
    If List1.ListIndex = -1 Then
    DBFire.Execute "UPDATE TB_WAS SET WAS_BOT_SEQ = NULL WHERE WAS_SEQ = " & WasSeq
    Else
    DBFire.Execute "UPDATE TB_WAS SET WAS_BOT_SEQ = " & GetSeq & " WHERE WAS_SEQ = " & WasSeq
    End If
End If
Unload Me
End Sub

Private Sub Form_Load()
RefreshList
UpdateLbl
End Sub

Private Sub List1_Click()
UpdateLbl
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
List1.ListIndex = -1
UpdateLbl
End If
End Sub

Private Sub Command1_Click()
List1.ListIndex = -1
UpdateLbl
End Sub

Private Sub UpdateLbl()
If List1.ListIndex = -1 Then
Label3 = Label3.Tag & "null"
Else
Label3 = Label3.Tag & GetSeq
End If
End Sub
