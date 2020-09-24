VERSION 5.00
Begin VB.Form frmEmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Empresa"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12270
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
   ScaleHeight     =   6705
   ScaleWidth      =   12270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Retornar"
      Height          =   360
      Left            =   180
      TabIndex        =   27
      Top             =   6240
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Enabled         =   0   'False
      Height          =   360
      Index           =   4
      Left            =   9780
      TabIndex        =   25
      Top             =   5700
      Width           =   450
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   360
      Index           =   4
      Left            =   11640
      TabIndex        =   24
      Top             =   5700
      Width           =   450
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Index           =   4
      Left            =   9780
      TabIndex        =   23
      Top             =   1500
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Index           =   3
      Left            =   7380
      TabIndex        =   21
      Top             =   1500
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   360
      Index           =   3
      Left            =   9240
      TabIndex        =   20
      Top             =   5700
      Width           =   450
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Enabled         =   0   'False
      Height          =   360
      Index           =   3
      Left            =   7380
      TabIndex        =   19
      Top             =   5700
      Width           =   450
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ok"
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   11760
      TabIndex        =   18
      Top             =   900
      Width           =   330
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ok"
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   11760
      TabIndex        =   17
      Top             =   360
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   180
      MaxLength       =   120
      TabIndex        =   15
      Top             =   900
      Width           =   11475
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   180
      MaxLength       =   120
      TabIndex        =   13
      Top             =   360
      Width           =   11475
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   4980
      TabIndex        =   11
      Top             =   5700
      Width           =   450
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   6840
      TabIndex        =   10
      Top             =   5700
      Width           =   450
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Index           =   2
      Left            =   4980
      TabIndex        =   8
      Top             =   1500
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   2580
      TabIndex        =   7
      Top             =   5700
      Width           =   450
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   360
      Index           =   1
      Left            =   4440
      TabIndex        =   6
      Top             =   5700
      Width           =   450
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Index           =   1
      Left            =   2580
      TabIndex        =   4
      Top             =   1500
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   5700
      Width           =   450
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   360
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   5700
      Width           =   450
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   1500
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numeros de Whatsapp"
      Height          =   195
      Index           =   4
      Left            =   9840
      TabIndex        =   26
      Top             =   1260
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bots"
      Height          =   195
      Index           =   3
      Left            =   7440
      TabIndex        =   22
      Top             =   1260
      Width           =   315
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seq:"
      Height          =   195
      Left            =   11640
      TabIndex        =   16
      Top             =   120
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição da Empresa:"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   14
      Top             =   660
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome da Empresa:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   12
      Top             =   120
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vinculos: [DEP : ATD]"
      Height          =   195
      Index           =   2
      Left            =   5040
      TabIndex        =   9
      Top             =   1260
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Departamentos:"
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   5
      Top             =   1260
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atendentes:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1260
      Width           =   900
   End
End
Attribute VB_Name = "frmEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal Millis As Long)

Dim Empresa As Long

Private Property Get GetSeq(ByVal List As Long, Optional ByVal Index As Long = -1) As Long
Dim C() As String
If Index < 0 Then Index = List1(List).ListIndex
If Index < 0 Then Exit Property
C = Split(List1(List).List(Index), ":")
On Error Resume Next
GetSeq = C(0)
End Property

Private Property Get GetCap(ByVal List As Long, Optional ByVal Index As Long = -1) As String
Dim C() As String
If Index < 0 Then Index = List1(List).ListIndex
If Index < 0 Then Exit Property
C = Split(List1(List).List(Index), ":")
On Error Resume Next
GetCap = C(1)
End Property

Private Sub RefreshList(ByVal Index As Long)
Dim F As New hDirectFire, Z As Long
Dim S As String, C(1) As String
Dim OldListIndex As Long
OldListIndex = List1(Index).ListIndex
List1(Index).Clear
List1(Index).Refresh
Select Case Index
Case 0: S = "SELECT ATD_SEQ,ATD_NAME FROM TB_ATD WHERE ATD_EMP_SEQ = " & Empresa & " ORDER BY ATD_SEQ": C(0) = "ATD_SEQ": C(1) = "ATD_NAME"
Case 1: S = "SELECT DEP_SEQ,DEP_NAME FROM TB_DEP WHERE DEP_EMP_SEQ = " & Empresa & " ORDER BY DEP_SEQ": C(0) = "DEP_SEQ": C(1) = "DEP_NAME"
Case 2: S = "SELECT VNC_SEQ,VNC_DEP_SEQ,VNC_ATD_SEQ FROM TB_VNC WHERE VNC_EMP_SEQ = " & Empresa & " ORDER BY VNC_SEQ"
Case 3: S = "SELECT BOT_SEQ,BOT_STATUS FROM TB_BOT WHERE BOT_EMP_SEQ = " & Empresa & " ORDER BY BOT_SEQ"
Case 4: S = "SELECT WAS_SEQ,WAS_BOT_SEQ,WAS_VAL FROM TB_WAS WHERE WAS_EMP_SEQ = " & Empresa & " ORDER BY WAS_SEQ"
End Select
If F.OpenRecordset(S) Then
    If Index = 2 Then
        Do Until F.EOF
        List1(Index).AddItem F("VNC_SEQ") & ":[" & F("VNC_DEP_SEQ") & "," & F("VNC_ATD_SEQ") & "]"
        F.MoveNext
        Loop
    ElseIf Index = 3 Then
        Do Until F.EOF
            Select Case F("BOT_STATUS")
            Case 0: List1(Index).AddItem F("BOT_SEQ") & ":Off"
            Case 1: List1(Index).AddItem F("BOT_SEQ") & ":On"
            Case 2: List1(Index).AddItem F("BOT_SEQ") & ":Reiniciando..."
            Case Else
            List1(Index).AddItem F("BOT_SEQ") & ":?(" & CStr(F("BOT_STATUS")) & ")"
            End Select
        F.MoveNext
        Loop
    ElseIf Index = 4 Then
        Do Until F.EOF
            If F("WAS_BOT_SEQ") Then
            List1(Index).AddItem F("WAS_SEQ") & ":" & F("WAS_BOT_SEQ") & ":" & F("WAS_VAL")
            Else
            List1(Index).AddItem F("WAS_SEQ") & ":null:" & F("WAS_VAL")
            End If
        F.MoveNext
        Loop
    Else
        Do Until F.EOF
        List1(Index).AddItem F(C(0)) & ":" & F(C(1))
        F.MoveNext
        Loop
    End If
Else
MsgBox F.LastErrDescription
End If
Sleep 20
If OldListIndex >= List1(Index).ListCount Then OldListIndex = OldListIndex - 1
List1(Index).ListIndex = OldListIndex
Command2(Index).Enabled = OldListIndex >= 0
End Sub

Private Sub Command3_Click(Index As Integer)
If Index = 0 Then
DBFire.Execute "UPDATE TB_EMP SET EMP_NAME = '" & Text1(0) & "' WHERE EMP_SEQ = " & Empresa
ElseIf Index = 1 Then
DBFire.Execute "UPDATE TB_EMP SET EMP_DESC = '" & Text1(1) & "' WHERE EMP_SEQ = " & Empresa
End If
Command3(Index).Enabled = False
End Sub

Private Sub Command4_Click()
frmMain.Show
Unload Me
End Sub

Private Sub Form_Load()
Visible = False
End Sub

Public Sub Present(ByVal Emp As Long)
Dim F As New hDirectFire
Empresa = Emp
Label3 = "Seq:" & Empresa
If F.OpenRecordset("SELECT EMP_SEQ,EMP_NAME,EMP_DESC FROM TB_EMP WHERE EMP_SEQ = " & Empresa & " ORDER BY EMP_SEQ") Then
    Text1(0) = F("EMP_NAME")
    Text1(1) = F("EMP_DESC")
    Command3(0).Enabled = False
    Command3(1).Enabled = False
End If
RefreshList 0
RefreshList 1
RefreshList 2
RefreshList 3
RefreshList 4
End Sub

Private Sub List1_Click(Index As Integer)
Command2(Index).Enabled = True
Command1(2).Enabled = Command2(0).Enabled And Command2(1).Enabled
End Sub

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Debug.Print KeyCode
If KeyCode = 116 Or KeyCode = 82 Then
RefreshList Index
End If
End Sub

Private Sub List1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then List1(Index).ListIndex = -1
Command1(2).Enabled = False
End Sub

Private Sub Text1_Change(Index As Integer)
Command3(Index).Enabled = True
End Sub

Private Sub Command1_Click(Index As Integer) '+ button
If Index = 0 Then
DBFire.Execute "INSERT INTO TB_ATD (ATD_SEQ,ATD_EMP_SEQ,ATD_NAME) VALUES (GEN_ID(GE_ATD_SEQ,1)," & Empresa & ",'Novo Atendente')"
ElseIf Index = 1 Then
DBFire.Execute "INSERT INTO TB_DEP (DEP_SEQ,DEP_EMP_SEQ,DEP_NAME) VALUES (GEN_ID(GE_DEP_SEQ,1)," & Empresa & ",'Novo Departamento')"
ElseIf Index = 2 Then
DBFire.Execute "INSERT INTO TB_VNC (VNC_SEQ,VNC_EMP_SEQ,VNC_DEP_SEQ,VNC_ATD_SEQ) VALUES (GEN_ID(GE_VNC_SEQ,1)," & Empresa & "," & GetSeq(1) & "," & GetSeq(0) & ")"
ElseIf Index = 3 Then
DBFire.Execute "INSERT INTO TB_BOT (BOT_SEQ,BOT_EMP_SEQ,BOT_STATUS) VALUES (GEN_ID(GE_BOT_SEQ,1)," & Empresa & ",0)"
ElseIf Index = 4 Then
    If GetSeq(3) = 0 Then
    DBFire.Execute "INSERT INTO TB_WAS (WAS_SEQ,WAS_EMP_SEQ,WAS_BOT_SEQ,WAS_VAL) VALUES (GEN_ID(GE_WAS_SEQ,1)," & Empresa & ",null,0)"
    Else
    DBFire.Execute "INSERT INTO TB_WAS (WAS_SEQ,WAS_EMP_SEQ,WAS_BOT_SEQ,WAS_VAL) VALUES (GEN_ID(GE_WAS_SEQ,1)," & Empresa & "," & GetSeq(3) & ",0)"
    End If
End If
RefreshList Index
End Sub

Private Sub Command2_Click(Index As Integer)
If Index = 0 Then
DBFire.Execute "DELETE FROM TB_ATD WHERE ATD_SEQ = " & GetSeq(0): RefreshList 2
ElseIf Index = 1 Then
DBFire.Execute "DELETE FROM TB_DEP WHERE DEP_SEQ = " & GetSeq(1): RefreshList 2
ElseIf Index = 2 Then
DBFire.Execute "DELETE FROM TB_VNC WHERE VNC_SEQ = " & GetSeq(2)
ElseIf Index = 3 Then
DBFire.Execute "DELETE FROM TB_BOT WHERE BOT_SEQ = " & GetSeq(3): RefreshList 4
ElseIf Index = 4 Then
DBFire.Execute "DELETE FROM TB_WAS WHERE WAS_SEQ = " & GetSeq(4)
End If
RefreshList Index
End Sub

Private Sub List1_DblClick(Index As Integer) 'open item
Dim S As String
If Index = 0 Then 'Atendentes
    S = GetCap(0)
RetryAtd:
    S = InputBox("Insira um Novo Nome Para Este Atendente", "Novo Valor", S)
    If LenB(S) = 0 Then Exit Sub
    If Len(S) > 120 Then
    MsgBox "Nome Muito Longo! É Possivel No Máximo de 120 Caracteres!", vbInformation, "Erro"
    GoTo RetryAtd
    End If
    DBFire.Execute "UPDATE TB_ATD SET ATD_NAME = '" & S & "' WHERE ATD_SEQ = " & GetSeq(0)
ElseIf Index = 1 Then 'Departamentos
    S = GetCap(1)
RetryDep:
    S = InputBox("Insira um Novo Nome Para Este Departamento", "Novo Valor", S)
    If LenB(S) = 0 Then Exit Sub
    If Len(S) > 120 Then
    MsgBox "Nome Muito Longo! É Possivel No Máximo de 120 Caracteres!", vbInformation, "Erro"
    GoTo RetryDep
    End If
    DBFire.Execute "UPDATE TB_DEP SET DEP_NAME = '" & S & "' WHERE DEP_SEQ = " & GetSeq(1)
ElseIf Index = 2 Then 'Vinculos
    MsgBox "Vinculos Não Sao Editaveis", vbInformation, ""
    Exit Sub
ElseIf Index = 3 Then 'Bots
    frmBot.BotSeq = GetSeq(3)
    frmBot.Show 1, Me
ElseIf Index = 4 Then 'Was
    frmWas.Empresa = Empresa
    frmWas.WasSeq = GetSeq(4)
    frmWas.Show 1, Me
End If
RefreshList CLng(Index)
End Sub


