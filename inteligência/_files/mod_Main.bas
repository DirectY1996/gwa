Attribute VB_Name = "mod_main"
Option Explicit

Public m_Resource As New hResource

Sub Main()
Dim cc As Long, mFile As String
Dim m_path_dll As String

    Dim m_ExeName As String
    m_ExeName = App.Path & "\" & App.EXEName
    If Not Right(LCase(m_ExeName), 4) = ".exe" Then
        m_ExeName = m_ExeName & ".exe"
    End If
    m_Resource.AssetsPath = App.Path & "\assets"
    m_Resource.ResourceFile = m_ExeName
    m_Resource.UseAssetsPath = IsIDE
    
    
    LoadOCX App.Path & "\codebank.ini", App.Title
    
    Dim txt As String, m_FileOut As String
    m_FileOut = App.Path & "\..\dados.fdb"
    If Not File.FileExist(m_FileOut) Then
        File.SetFileBytesEx m_FileOut, m_Resource.LoadResource("def-v1.fdb")
    End If
    If File.FileExist(m_FileOut) Then
        If OpenDB(m_FileOut) Then
            frmMain.Show
        Else
            MsgBox "Não abriu o banco de dados"
        End If
    Else
        MsgBox "Arquivo não existe"
    End If
End Sub

Public Function IsIDE() As Boolean
Static mIsIDE As Boolean
    Debug.Assert SetTrue(mIsIDE) = False
    IsIDE = mIsIDE
End Function

Function SetTrue(mBool As Boolean) As Boolean
    mBool = True
End Function
