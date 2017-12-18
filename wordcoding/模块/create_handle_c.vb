Module create_handle_c


Sub create_handle_c(ByVal value)

    Dim i As Integer
    Dim j As Integer
    Dim print_table As single_table
    Dim SQLStr As String
    Dim program As New program_class
    Dim comments As New comments_class
    Dim single_file As in_file
    Dim str As String
    Dim handle_file As Object
    Dim handle_file_txt As String
    Dim s_file As Object
    Dim ts As Object, s As String
    Dim tmp_str As String
    Dim flag As Integer


        Dim docpath As Object
        docpath = env_conf.win_work_path & "\" & "src\"
        program = value
        If Dir(docpath, vbDirectory) = "" Then
            MkDir(docpath)
        End If
        FileOpen(1, docpath & program.program_english_name & "_handle" & ".c", OpenMode.Output)

        Print(1, "#include """ & program.program_english_name & ".h""" & vbCrLf)
        Print(1, "#include ""endday_db_diy.h""" & vbCrLf)
        Print(1, "" & vbCrLf)
        SQLStr = UCase(program.program_english_name) & "_HANDLE_ARGV_IN_REC *" & Space(1) & "in_set," & UCase(program.program_english_name) & "_HANDLE_ARGV_OUT_REC *" & Space(1) & "out_set"
        '        For i = 1 To program.program_cust_handle.func_argv_class.Count
        '            SQLStr = SQLStr & program.program_cust_handle.func_argv_class.Item(i)
        '            If i < program.program_cust_handle.func_argv.Count Then
        '                SQLStr = SQLStr & ","
        '            End If
        '        Next i
        'MsgBox(SQLStr)
        Print(1, "int " & program.program_english_name & "_handle(" & SQLStr & ")" & vbCrLf)
        Print(1, "{" & vbCrLf)

        '编写注释说明
        Print(1, Chr(9) & "/*系统自动补充代码功能接口说明*/" & vbCrLf)
        Print(1, "" & vbCrLf)


        Print(1, Chr(9) & "int i=0;" & vbCrLf)
        Print(1, Chr(9) & "int iRet=0;" & vbCrLf)

        SQLStr = ""
        For i = 1 To program.program_cust_handle.func_argv_class.Count
            Print(1, "/*" & vbCrLf)
            comments = program.program_cust_handle.argv_comments.Item(i)
            SQLStr = Chr(9) & "//" & program.program_cust_handle.func_argv_class.Item(i)
            Print(1, SQLStr & vbCrLf)
            Print(1, Chr(9) & "----{" & vbCrLf)
            Print(1, Chr(9) & Chr(9) & "|" & vbCrLf)
            For j = 1 To comments.str_type.Count
                Print(1, Chr(9) & Chr(9) & "|" & comments.str_type.Item(j) & Chr(9) & comments.str_def.Item(j) & Chr(9) & comments.str_note.Item(j) & vbCrLf)
            Next j
            Print(1, Chr(9) & "----}" & vbCrLf)
            Print(1, "*/" & vbCrLf)
        Next i

        Print(1, "" & vbCrLf)
        Print(1, Chr(9) & "/*系统自动补充代码功能接口说明结束*/" & vbCrLf)
        Print(1, "" & vbCrLf)
        'Print #1, Chr(9) & "strcpy(" & single_file.file_heads_val & "->total_count, ""100"");"




        '        For i = 1 To program.in_file.Count
        '           single_file = program.in_file.Item(i)
        '            Print #1, Chr(9) & "for(i=0;i<atoi(in_set->" & single_file.file_heads_val & ".total_count);i++)"
        '            Print #1, Chr(9) & "{"
        '            For j = 1 To single_file.file_bodys.str_name.Count
        '                str = single_file.file_bodys.str_type.Item(j)
        '                If InStr(str, "double") > 0 Or InStr(str, "int") > 0 Then
        '                    SQLStr = "LOG_INFO(""p_" & single_file.file_bodys_val & "." & single_file.file_bodys.str_name.Item(j) & "=[%.2f]""," & "in_set->" & "p_" & single_file.file_bodys_val & "[i]." & single_file.file_bodys.str_name.Item(j) & ");"
        '                Else
        '                    SQLStr = "LOG_INFO(""p_" & single_file.file_bodys_val & "." & single_file.file_bodys.str_name.Item(j) & "=[%s]""," & "in_set->" & "p_" & single_file.file_bodys_val & "[i]." & single_file.file_bodys.str_name.Item(j) & ");"
        '                End If
        '                Print #1, Chr(9) & Chr(9) & SQLStr
        '            Next j
        '            Print #1, Chr(9) & Chr(9) & "LOG_INFO("" [%d]"",i);"
        '            Print #1, Chr(9) & "}"
        '
        '        Next i
        '        tmp_str = "in_set->" & single_file.file_heads_val & ".total_count"


        '        For i = 1 To program.out_file.Count
        '           single_file = program.out_file.Item(i)
        '
        '            Print #1, Chr(9) & "strcpy(out_set->" & single_file.file_heads_val & ".total_count, " & tmp_str & ");"
        '            Print #1, Chr(9) & UCase(single_file.file_name) & "_BODY_REC " & Space(1) & single_file.file_bodys_val & "_1;"
        '            Print #1, Chr(9) & "out_set->" & "p_" & single_file.file_bodys_val & " = (" & UCase(single_file.file_name) & "_BODY_REC *)calloc(100," & UCase(single_file.file_name) & "_BODY_REC_LEN" & ");"
        '            Print #1, Chr(9) & "for(i=0;i<atoi(out_set->" & single_file.file_heads_val & ".total_count);i++)"
        '            Print #1, Chr(9) & "{"
        '            Print #1, Chr(9) & Chr(9) & "memset(&" & single_file.file_bodys_val & "_1, 0x00," & UCase(single_file.file_name) & "_BODY_REC_LEN);"
        '            For j = 1 To single_file.file_bodys.str_name.Count
        '                str = single_file.file_bodys.str_type.Item(j)
        '                If InStr(str, "double") > 0 Or InStr(str, "int") > 0 Then
        '                    Print #1, Chr(9) & Chr(9) & single_file.file_bodys_val & "_1." & single_file.file_bodys.str_name.Item(j) & " = 11;"
        '                Else
        '                    Print #1, Chr(9) & Chr(9) & "strcpy(" & single_file.file_bodys_val & "_1." & single_file.file_bodys.str_name.Item(j) & ", ""11"");"
        '                End If
        '
        '            Next j
        '            Print #1, Chr(9) & Chr(9) & "LOG_INFO("" [%d]"",i);"
        '            Print #1, Chr(9) & Chr(9) & "out_set->" & "p_" & single_file.file_bodys_val & "[i] = " & single_file.file_bodys_val & "_1;"
        '            Print #1, Chr(9) & Chr(9) & "LOG_INFO("" [%s]""," & "out_set->" & "p_" & single_file.file_bodys_val & "[i].currency);"
        '            Print #1, Chr(9) & "}"
        '
        '            For j = 1 To single_file.file_heads.str_name.Count
        '                str = single_file.file_heads.str_type.Item(j)
        '                If InStr(str, "double") > 0 Or InStr(str, "int") > 0 Then
        '                    Print #1, Chr(9) & Chr(9) & "out_set->" & single_file.file_heads_val & "." & single_file.file_heads.str_name.Item(j) & " = 100;"
        '                Else
        '                    Print #1, Chr(9) & Chr(9) & "strcpy(" & "out_set->" & single_file.file_heads_val & "." & single_file.file_heads.str_name.Item(j) & ", ""100"");"
        '                End If
        '
        '            Next j
        '        Next i


        Print(1, "" & vbCrLf)

        '        For i = 1 To 10
        '            Print #1, Chr(10)
        '        Next i
        REM 打开远程下载到的文件

        handle_file = CreateObject("Scripting.FileSystemObject")
        On Error GoTo coin_bow

        s_file = handle_file.GetFile(docpath & "\src\remote_host\" & program.program_english_name & "_handle" & ".c")
        MsgBox(s_file)
        ts = s_file.OpenAsTextStream(1, -2)
        
        flag = 0

        Do While Not ts.AtEndOfStream
            s = ""
            If flag <> 1 Then
                s = ts.ReadLine
                MsgBox(s)
            End If
            REM MsgBox s

            If InStr(s, "/*coin_bow-wow the following regional user custom code*/") > 0 Then
                '                MsgBox s
                REM 定位到客户自定义代码的开始，copy
                '                        s = ts.ReadLine
                flag = 0
                Do Until ts.AtEndOfStream

                    REM copy to 1
                    If InStr(s, "coin_bow-wow to this end") > 0 Then
                        'MsgBox s
                        'MsgBox ts.Line
                        '                                    Print #1, s
                        flag = 1
                        Exit Do
                    Else
                        Print(1, Replace(s, Chr(13) & Chr(7) & Chr(10) & vbCrLf, "") & vbCrLf)
                    End If
                    Print(2, s & vbCrLf)
                    s = ts.ReadLine
                Loop
            End If
            If InStr(s, "coin_bow-wow to this end") > 0 Then
                '                        s = ts.ReadLine
                flag = 0
                'MsgBox ts.Line
                Do Until ts.AtEndOfStream
                    'MsgBox ts.Line

                    REM copy to 1
                    If InStr(s, "coin_bow-wow the following regional user custom code") > 0 Then
                        '                                    MsgBox s
                        '                                    Print #1, s
                        flag = 1
                        Exit Do
                    Else
                        Print(1, Replace(s, Chr(13) & Chr(7) & Chr(10) & vbCrLf, "") & vbCrLf)
                    End If
                    Print(2, s & vbCrLf)
                    s = ts.ReadLine
                Loop

            End If
        Loop
        ts.Close
        Print(1, "" & vbCrLf)

        FileClose(1)
        Exit Sub
coin_bow:
        Print(1, Chr(9) & "/*coin_bow-wow the following regional user custom code*/" & vbCrLf)
        Print(1, Chr(9) & "/*coin_bow-wow to this end*/" & vbCrLf)
        Print(1, Chr(9) & "return 0;")
        Print(1, "}" & vbCrLf)
        Print(1, "" & vbCrLf)
        Print(1, "" & vbCrLf)
        Print(1, "" & vbCrLf)
        Print(1, "" & vbCrLf)
        Print(1, "" & vbCrLf)
        Print(1, "" & vbCrLf)
        FileClose(1)
    End Sub







End Module
