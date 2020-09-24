Attribute VB_Name = "mod_update"
Option Explicit

Dim m_key_step As String
Dim m_progress As hProgress
Dim dbStep_total As Long

Function GetDBKeys() As String()
Dim F As Long, S As String
F = FreeFile
S = String$(FileLen(App.Path & "\Keys.txt"), 0)
Open App.Path & "\Keys.txt" For Binary As F
Get F, , S
Close F
GetDBKeys = Split(S, "|")
End Function

Function OpenDB(ByVal m_DataFile As String) As Boolean
Dim m_ConnectionString As New hStringBuilder
Dim C() As String
        
        
        'sql = sql & "Client="
        'sql = sql & App.Path & "\fbclient255.dll;"
        
        'sql = sql & "\Firebird-2.5.8.27089-0_Win32_embed\fbembed.dll;"
        'sql = sql & "R:\Firebird-3.0.3.32900-0_Win32\fbclient.dll"
        'sql = sql & "R:\Firebird-2.5.8.27089-0_x64_embed\fbembed.dll;"
        
   On Error GoTo OpenDB_Error
        
        C = GetDBKeys
        
        m_ConnectionString.Append "DRIVER=Firebird/InterBase(r) driver; UID=" & C(0) & "; PWD=" & C(1) & "; "
        m_ConnectionString.Append "DBNAME=localhost/19259:" & m_DataFile & ";"
        m_ConnectionString.Append "Client=" & App.Path & "\lib\fbclient259.dll;"
        m_ConnectionString.Append ";CHARSET=win1252;"
        
        OpenDB = DBFire.OpenData(m_ConnectionString.toString, C(0), C(1))
        If OpenDB Then
            If DBFire.CountSQL("SELECT a.RDB$RELATION_NAME FROM RDB$RELATIONS a WHERE RDB$SYSTEM_FLAG = 0 AND RDB$RELATION_TYPE = 0") = 0 Then
                Update_initial
            End If
            Update_domains
            Update_tables_main
            
            
            If Not m_progress Is Nothing Then
                m_progress.EndBegin
                Set m_progress = Nothing
            End If
        Else
            MsgBox DBFire.LastOpenError
        End If
   On Error GoTo 0
   Exit Function

OpenDB_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure OpenDB of Módulo mod_update, " & Erl
End Function

Public Function GetSYSLong(ByVal SysID As String, Optional mDefValue As Long) As Long
    Dim tabela As New hDirectFire
    On Error GoTo GetSYSLong_Error
    If tabela.CreateQuery("select SYS_INT from TB_SYS_INFO where SYS_KEY = '" & SysID & "'") Then
        If tabela.OpenRecordset Then
            If tabela.RegSelected Then
                GetSYSLong = tabela("SYS_INT")
            Else
                GetSYSLong = mDefValue
                SetSYSLong SysID, mDefValue
            End If
            tabela.CloseTable
        End If
    End If
    On Error GoTo 0
    Exit Function
GetSYSLong_Error:
    MsgBox Err.Description, , "GetSYSLong"
    'ReportErr

End Function

Public Sub SetSYSLong(ByVal SysID As String, ByVal mData As Long)
    Dim tabela As New hDirectFire
    Dim sql As String
    sql = "select SYS_KEY,SYS_INT, SYS_SEQ from TB_SYS_INFO where SYS_KEY = '" & SysID & "'"
    If tabela.OpenRecordset(sql) Then
        If tabela.RegSelected Then
            tabela.Edit
        Else
            tabela.AddNew
            tabela("SYS_SEQ") = DBFire.GetGenerator("GE_SYS_SEQ")
            tabela("SYS_KEY") = SysID
        End If
        tabela("SYS_INT") = mData
        tabela.Update
        tabela.CloseTable
    End If
End Sub

Sub Update_initial()
'Dim m_sql As String
'Dim dbStep  As Long
'    dbStep_total = -1
End Sub

Sub Update_domains()
'Dim dbStep  As Long
'    m_key_step = "bd_update_domains"
'    dbStep = GetSYSLong(m_key_step)
'    dbStep_total = -1
End Sub

Sub Update_tables_main()
'Dim dbStep  As Long
'Dim m_txt As String
'    m_key_step = "bd_update_tables_main"
'    dbStep = GetSYSLong(m_key_step)
'    dbStep_total = -1
End Sub

Sub CloseDB()
    If DBFire.IsOpen Then
        DBFire.CloseData
    End If
End Sub

Sub update_1()
Dim dbStep As Long
End Sub

Function Exec(dbStep As Long, ByVal txt As String, Optional IgnoreError As Boolean) As Boolean
    Dim tabela As New hDirectFire
        If (dbStep_total - dbStep) > 10 Then
            If m_progress Is Nothing Then
                Set m_progress = New hProgress
                m_progress.Begin
            End If
            m_progress.SetProgress dbStep, dbStep_total
        End If
        '    'If m_flag_1 Then
        '    '    Exec = True
        '    '    Exit Function
        '    'End If
        '    'frmMainProgress.cProgress1.Value = (dbStep / m_total) * 100
        '    'frmMainMDI.AddText "Atualizando... " & txt
        '    If LCase(txt) = "nop" Then
        '        dbStep = dbStep + 1
        '        Exec = True
        '        Call SetSYSLong("bd_update", dbStep)
        '        Exit Function
        '    End If
        '    If txt = "NOPREC" Then
        '        dbStep = dbStep + 1
        '        Exec = True
        '        Call SetSYSLong("bd_update", dbStep)
        '        Exit Function
        '    End If
            'IgnoreError = True
   On Error GoTo Exec_Error

    DBFire.BeginTrans
    If tabela.CreateQuery(txt) Then
        Exec = tabela.Execute
        If Exec Then
            tabela.CloseTable
        Else
            'frmMain.addlog "Erro em: " & m_key_step & " = " & dbStep
            'frmMain.addlog tabela.LastErrDescription
            
            Debug.Print "-----------------------------------------------------"
            Debug.Print "Erro em: " & m_key_step & " = " & dbStep
            Debug.Print tabela.LastErrDescription
            Debug.Print txt
            Debug.Print
            Debug.Print
        End If
    End If
    If Exec Then
        'GravaDBLog "OK - " & Date & ";" & Time & ";" & dbStep
        dbStep = dbStep + 1
        If Len(m_key_step) Then
            Call SetSYSLong(m_key_step, dbStep)
        End If
        DBFire.CommitTrans
    Else
        DBFire.Rollback
        '        'GravaDBLog "ER - " & Date & ";" & Time & ";" & dbStep & " " & Tabela.LastErrDescription
        '        If Not IgnoreError Then
        '            MsgBox txt
        '            MsgBox tabela.LastErrDescription
        '            Debug.Print tabela.LastErrDescription
        '        End If
        '
        '        If IgnoreError Then
        '            dbStep = dbStep + 1
        '            Call SetSYSLong("bd_update", dbStep)
        '        End If
    End If
    'SetCaption "Atualizando..." & CStr(dbStep) & " " & txt
    'System.Sleep 50

   On Error GoTo 0
   Exit Function

Exec_Error:
    DBFire.Rollback
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Exec of Módulo mod_update, " & Erl
End Function

