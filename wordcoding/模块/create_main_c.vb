Module create_main_c
    Sub create_mian_c(ByVal value)

        Dim i As Integer
        Dim SQLStr As String
        Dim program As New program_class
        Dim single_file As in_file
        Dim str As String
        Dim docpath As String = Nothing

        env_conf.win_sys_dir = "C:\Program Files (x86)\Microsoft\coding_create"


        docpath = env_conf.win_work_path & "\"
        '"E:\邮储理财项目组工作\新技术\VBA" & "\"
        program = value

        If Dir(docpath & "src\", vbDirectory) = "" Then
            MkDir(docpath & "src\")
        End If
        FileOpen(1, docpath & "src\" & program.program_english_name & "_main" & ".c", OpenMode.Output)

        Print(1, "#include """ & program.program_english_name & ".h""" & vbCrLf)
        Print(1, "#include <stdio.h>" & vbCrLf)
        Print(1, "#include <stdlib.h>" & vbCrLf)
        Print(1, "#include <string.h>" & vbCrLf)
        Print(1, "#include <pubhead.h>" & vbCrLf)
        Print(1, "" & vbCrLf)

        REM 创建客户自定义函数的形参结构体



        Print(1, "" & vbCrLf)
        Print(1, "int main(int argc,char ** argv)" & vbCrLf)
        Print(1, "{" & vbCrLf)
        Print(1, Chr(9) & "int iRet=0;" & vbCrLf)
        Print(1, Chr(9) & "char logfile[128] = """";" & vbCrLf)
        Print(1, Chr(9) & "char ch;" & vbCrLf)
        '        Print #1, Chr(9) & "SYS_DATETIME sys_date; " & Space(25) & "/* 公共参数表信息 */"



        Print(1, Chr(9) & UCase(program.program_english_name) & "_HANDLE_ARGV_IN_REC" & Space(1) & "in_set;" & vbCrLf)
        Print(1, Chr(9) & UCase(program.program_english_name) & "_HANDLE_ARGV_OUT_REC" & Space(1) & "out_set;" & vbCrLf)

        Print(1, Chr(9) & program.call_method_struct_val & ";" & vbCrLf)



        '        Print #1, Chr(9) & "memset(&sys_date, 0x00, SYS_DATETIME_LEN);"









        '        For i = 1 To program.in_file.Count
        '           single_file = New in_file
        '           single_file = program.in_file.Item(i)
        '            Print #1, ""
        '            Print #1, Chr(9) & single_file.file_formats_struct_init & ";"
        '            Print #1, Chr(9) & single_file.file_heads_struct_init & ";"
        '        Next i
        '
        '        Rem 输出文件
        '        For i = 1 To program.out_file.Count
        '           single_file = New in_file
        '           single_file = program.out_file.Item(i)
        '            Print #1, ""
        '            Print #1, Chr(9) & single_file.file_formats_struct_init & ";"
        '            Print #1, Chr(9) & single_file.file_heads_struct_init & ";"
        '        Next i


        Print(1, Chr(9) & program.call_method_struct_val_init & ";" & vbCrLf)
        Print(1, Chr(9) & "memset(&in_set" & " ,0X00,sizeof(" & UCase(program.program_english_name) & "_HANDLE_ARGV_IN_REC));" & vbCrLf)
        Print(1, Chr(9) & "memset(&out_set" & " ,0X00,sizeof(" & UCase(program.program_english_name) & "_HANDLE_ARGV_OUT_REC));" & vbCrLf)
        Print(1, "" & vbCrLf)



        'Print(1, Chr(9) & "If (argc < " & program.call_method.argv_flag.Count + 1 & ")" & vbCrLf)
        'Print(1, Chr(9) & "{" & vbCrLf)
        'str = program.program_english_name
        'For i = 1 To program.call_method.argv_flag.Count
        '    str = str & Space(1) & "-" & program.call_method.argv_flag.Item(i) & program.call_method.argv_note(i)
        'Next i
        'Print(1, Chr(9) & Chr(9) & "fprintf(stderr,""" & str & "\n"");" & vbCrLf)
        'Print(1, Chr(9) & Chr(9) & "Return -1;" & vbCrLf)
        'Print(1, Chr(9) & "}" & vbCrLf)
        Print(1, Chr(9) & "" & vbCrLf)

        str = ""
        For i = 1 To program.call_method.argv_flag.Count
            If program.call_method.val_type_pre.Item(i) = "bool" Then
                str = str & program.call_method.argv_flag.Item(i) & ""
            Else
                str = str & program.call_method.argv_flag.Item(i) & ":"
            End If
        Next i
        Print(1, Chr(9) & "while ((ch = getopt(argc, argv, """ & str & """)) != -1)" & vbCrLf)
        Print(1, Chr(9) & "{" & vbCrLf)
        Print(1, Chr(9) & Chr(9) & "switch(ch)" & vbCrLf)
        Print(1, Chr(9) & Chr(9) & "{" & vbCrLf)
        For i = 1 To program.call_method.argv_flag.Count
            str = str & program.call_method.argv_flag.Item(i) & ""
            Print(1, Chr(9) & Chr(9) & Chr(9) & "case '" & program.call_method.argv_flag.Item(i) & "':" & vbCrLf)
            If program.call_method.val_type_pre.Item(i) = "bool" Then
            Else
                Print(1, Chr(9) & Chr(9) & Chr(9) & Chr(9) & "strcpy(in_set." & LCase(program.program_english_name) & "_argv_rec." & program.call_method.val_name.Item(i) & ",optarg);" & vbCrLf)
            End If

            Print(1, Chr(9) & Chr(9) & Chr(9) & Chr(9) & "in_set." & LCase(program.program_english_name) & "_argv_rec.index['" & program.call_method.argv_flag.Item(i) & "']=1;" & vbCrLf)
            Print(1, Chr(9) & Chr(9) & Chr(9) & Chr(9) & "break;" & vbCrLf)
        Next i
        Print(1, Chr(9) & Chr(9) & Chr(9) & "default:" & vbCrLf)
        Print(1, Chr(9) & Chr(9) & Chr(9) & Chr(9) & "break;" & vbCrLf)
        Print(1, Chr(9) & Chr(9) & "}" & vbCrLf)
        Print(1, Chr(9) & "}" & vbCrLf)

        ' 是否可选

        str = ""
        Dim o_flag As Integer
        o_flag = 0
        For i = 1 To program.call_method.argv_flag.Count
            MsgBox(LCase(program.program_english_name) & " " & program.call_method.argv_flag.Count)
            MsgBox(LCase(program.program_english_name) & " " & program.call_method.choose_flag.Item(i))
            If (program.call_method.choose_flag.Item(i) = "n") Then

                'str = str & "in_set." & LCase(program.program_english_name) & "_argv_rec.index['" & program.call_method.argv_flag.Item(i) & "']"
                'MsgBox(i & " " & str)
                If o_flag = 1 Then
                    str = str & " && " & "in_set." & LCase(program.program_english_name) & "_argv_rec.index['" & program.call_method.argv_flag.Item(i) & "']"
                Else
                    str = str & "in_set." & LCase(program.program_english_name) & "_argv_rec.index['" & program.call_method.argv_flag.Item(i) & "']"
                End If
                If i <> program.call_method.argv_flag.Count Then
                    o_flag = 1
                End If

            End If

            If i >= program.call_method.argv_flag.Count And o_flag = 1 Then
                str = "if(" & str & "!=1)"
                Print(1, Chr(9) & str & vbCrLf)
                Print(1, Chr(9) & "{" & vbCrLf)
                Print(1, Chr(9) & Chr(9) & "fprintf(stderr,""" & "必选参数不存在,请核实!" & "\n"");" & vbCrLf)
                Print(1, Chr(9) & Chr(9) & "return -1;" & vbCrLf)
                Print(1, Chr(9) & "}" & vbCrLf)
            End If
        Next i


        Print(1, Chr(9) & "" & vbCrLf)



        REM 文件变量定义
        For i = 1 To program.in_file.Count
            single_file = New in_file
            single_file = program.in_file.Item(i)
            Print(1, Chr(9) & "sprintf(in_set.file_name_" & i & ",""%s/%s/%s"", getenv(""SHARED_HOME""),""gf_sort_file"", """ & single_file.file_name & "." & single_file.file_formats.file_type & """); " & vbCrLf)
            Print(1, "" & vbCrLf)
        Next i

        REM 输出文件
        For i = 1 To program.out_file.Count
            single_file = New in_file
            single_file = program.out_file.Item(i)
            Print(1, Chr(9) & "sprintf(out_set.file_name_" & i & ",""%s/%s/%s"", getenv(""SHARED_HOME""), ""gf_sort_file"",""" & single_file.file_name & "." & single_file.file_formats.file_type & """); " & vbCrLf)
            Print(1, "" & vbCrLf)
        Next i

        '        For i = 1 To program.tables_struct_val.Count
        '            Print #1, Chr(9) & program.tables_struct_val_init.Item(i) & ";"
        '        Next i


        Print(1, Chr(9) & "/* main_code_begin */" & vbCrLf)


        Print(1, Chr(9) & "iRet = get_host_time(&in_set.sys_date);" & vbCrLf)
        Print(1, Chr(9) & "if (iRet != 0)" & vbCrLf)
        Print(1, Chr(9) & "{" & vbCrLf)
        Print(1, Chr(9) & Chr(9) & "LOG_ERR(""get_host_time error! i_rc[%d] "", iRet);" & vbCrLf)
        Print(1, Chr(9) & Chr(9) & "return iRet ;" & vbCrLf)
        Print(1, Chr(9) & "}" & vbCrLf)
        Print(1, Chr(9) & "" & vbCrLf)


        Print(1, Chr(9) & "sprintf(logfile, ""%s/log/errlog/%s_%s.log"", getenv(""HOME""), argv[0], in_set.sys_date.ca_date);" & vbCrLf)
        Print(1, Chr(9) & "log_set_filename(logfile);" & vbCrLf)
        Print(1, Chr(9) & "log_add_debugno();" & vbCrLf)
        Print(1, Chr(9) & "log_read_level(argv[0]);" & vbCrLf)
        Print(1, Chr(9) & "" & vbCrLf)


        Print(1, Chr(9) & "LOG_WARN("" 处理开始! "");" & vbCrLf)
        Print(1, Chr(9) & "LOG_INFO(""服务器机器日期[%s] "", in_set.sys_date.ca_date);" & vbCrLf)
        Print(1, Chr(9) & "LOG_INFO(""服务器机器时间[%s] "", in_set.sys_date.ca_time);" & vbCrLf)
        Print(1, Chr(9) & "LOG_INFO(""连接数据库!"");" & vbCrLf)
        Print(1, Chr(9) & "if ((iRet=db_open()) != 0)" & vbCrLf)
        Print(1, Chr(9) & "{" & vbCrLf)
        Print(1, Chr(9) & Chr(9) & "LOG_ERR(""db_open error sqlcode[%d] "", iRet);" & vbCrLf)
        Print(1, Chr(9) & Chr(9) & "fprintf(stderr, "" db_open; Error; sqlcode; [%d]\n"", iRet);" & vbCrLf)
        Print(1, Chr(9) & Chr(9) & "return -1;" & vbCrLf)
        Print(1, Chr(9) & "}" & vbCrLf)
        Print(1, Chr(9) & "" & vbCrLf)

        Print(1, Chr(9) & "strcpy(in_set.sys_date.st_para.organcode, ""11005293"");" & vbCrLf)
        Print(1, Chr(9) & "iRet = select_fd_publicpara(&in_set.sys_date.st_para);" & vbCrLf)
        Print(1, Chr(9) & "if (iRet != 0)" & vbCrLf)
        Print(1, Chr(9) & "{" & vbCrLf)
        Print(1, Chr(9) & Chr(9) & "LOG_ERR(""select_fd_publicpara error! i_rc[%d] "", iRet);" & vbCrLf)
        Print(1, Chr(9) & Chr(9) & "return iRet ;" & vbCrLf)
        Print(1, Chr(9) & "}" & vbCrLf)
        Print(1, Chr(9) & "" & vbCrLf)
        '        For i = 0 To program.tables.table_count - 1
        '           print_table = program.tables.get_single_table(i)
        '            For j = 1 To print_table.oper_type.Count
        '            Select Case print_table.oper_type(j)
        '                    Case "查询"
        '                        SQLStr = "iRet=" & "fetch" & "_" & LCase(print_table.table_name) & "_" & Format(CStr(j), "00") & "(" & program.tables_struct_fetch_argv.Item(i + 1) & ")" & ";"
        '                        Print #1, Chr(9) & SQLStr
        '                        Print #1, Chr(9) & "if(iRet!=0)"
        '                        Print #1, Chr(9) & "{"
        '                        Print #1, Chr(9) & Chr(9) & "printf(""" & SQLStr & "Error!" & "i_rc=[%d];"", iRet);"
        '                        Print #1, Chr(9) & Chr(9) & "return iRet;"
        '                        Print #1, Chr(9) & "}"
        '                        Print #1, ""
        '                    Case Else
        '                    End Select
        '
        '            Next j
        '            Print #1, ""
        '        Next i
        Print(1, "" & vbCrLf)
        Print(1, "" & vbCrLf)
        Print(1, Chr(9) & "//输入文件打开" & vbCrLf)
        For i = 1 To program.in_file.Count
            single_file = New in_file
            single_file = program.in_file.Item(i)
            Print(1, Chr(9) & "LOG_INFO(" & """文件[%s]打开准备打开操作!""," & "in_set.file_name_" & i & ");" & vbCrLf)
            '            Print #1, Chr(9) & "if ( ( p_in_file_" & i & " = fopen(""" & single_file.file_name & """, "" r"" ) ) == NULL )"
            '            Print #1, Chr(9) & "{"
            '            Print #1, Chr(9) & Chr(9) & "if(p_in_file_" & i & "!=NULL) fclose( " & "p_in_file_" & i & ");"
            '            Print #1, Chr(9) & Chr(9) & "return (-1);"
            '            Print #1, Chr(9) & "}"
            Print(1, "" & vbCrLf)
            Print(1, Chr(9) & "iRet=" & single_file.file_name & "_get( in_set.file_name_" & i & ",&in_set." & single_file.file_formats_val & " , &in_set." & single_file.file_heads_val & " , &in_set.p_" & single_file.file_bodys_val & ");" & vbCrLf)
            Print(1, Chr(9) & "if(iRet!=0)" & vbCrLf)
            Print(1, Chr(9) & "{" & vbCrLf)
            Print(1, Chr(9) & Chr(9) & "LOG_ERR(" & """文件[%s]打开失败!""," & "in_set.file_name_" & i & ");" & vbCrLf)
            Print(1, Chr(9) & Chr(9) & "return (iRet);" & vbCrLf)
            Print(1, Chr(9) & "}" & vbCrLf)
            Print(1, Chr(9) & "LOG_INFO(" & """文件[%s]输入结束!""," & "in_set.file_name_" & i & ");" & vbCrLf)
            Print(1, Chr(9) & "LOG_INFO(" & """文件[%s]输入结构体指针[%p]!""," & "in_set.file_name_" & i & ",in_set.p_" & single_file.file_bodys_val & ");" & vbCrLf)
            Print(1, "" & vbCrLf)

        Next i


        Print(1, "" & vbCrLf)
        Print(1, "" & vbCrLf)
        Print(1, Chr(9) & "//用户自定义处理流程开始" & vbCrLf)
        SQLStr = "&in_set , &out_set"
        '        For i = 1 To program.program_cust_handle.func_argv.Count
        '            SQLStr = SQLStr & "&" & program.program_cust_handle.func_argv.Item(i)
        '            If i < program.program_cust_handle.func_argv.Count Then
        '                SQLStr = SQLStr & ","
        '            End If
        '        Next i

        SQLStr = "iRet=" & program.program_english_name & "_handle(" & SQLStr & ");"
        Print(1, Chr(9) & SQLStr & vbCrLf)
        Print(1, Chr(9) & "if(iRet!=0)" & vbCrLf)
        Print(1, Chr(9) & "{" & vbCrLf)
        Print(1, Chr(9) & Chr(9) & "LOG_ERR(""" & SQLStr & "Error!" & "i_rc=[%d];"", iRet);" & vbCrLf)
        Print(1, Chr(9) & Chr(9) & "return iRet;" & vbCrLf)
        Print(1, Chr(9) & "}" & vbCrLf)
        Print(1, "" & vbCrLf)

        Print(1, Chr(9) & "/*用户自定义处理流程结束*/" & vbCrLf)


        Print(1, "" & vbCrLf)
        Print(1, "" & vbCrLf)
        Print(1, Chr(9) & "//输出文件写入" & vbCrLf)
        For i = 1 To program.out_file.Count
            single_file = New in_file
            single_file = program.out_file.Item(i)
            '            Print #1, Chr(9) & "if ( ( p_out_file_" & i & " = fopen(""" & single_file.file_name & """, "" w"" ) ) == NULL )"
            '            Print #1, Chr(9) & "{"
            '            Print #1, Chr(9) & Chr(9) & "if(p_out_file_" & i & "!=NULL) fclose( " & "p_out_file_" & i & ");"
            '            Print #1, Chr(9) & Chr(9) & "return (-1);"
            '            Print #1, Chr(9) & "}"
            Print(1, "" & vbCrLf)
            Print(1, Chr(9) & "iRet=" & single_file.file_name & "_put( out_set.file_name_" & i & ",&out_set." & single_file.file_formats_val & " , &out_set." & single_file.file_heads_val & " , out_set.p_" & single_file.file_bodys_val & ");" & vbCrLf)
            Print(1, Chr(9) & "if(iRet!=0)" & vbCrLf)
            Print(1, Chr(9) & "{" & vbCrLf)
            Print(1, Chr(9) & Chr(9) & "LOG_ERR(" & """文件[%s]输出失败!""," & "out_set.file_name_" & i & ");" & vbCrLf)
            Print(1, Chr(9) & Chr(9) & "return (iRet);" & vbCrLf)
            Print(1, Chr(9) & "}" & vbCrLf)
            Print(1, "" & vbCrLf)
        Next i

        Print(1, "" & vbCrLf)
        Print(1, "" & vbCrLf)
        '        For i = 0 To program.tables.table_count - 1
        '           print_table = program.tables.get_single_table(i)
        '            For j = 1 To print_table.oper_type.Count
        '            Select Case print_table.oper_type(j)
        '                    Case "更新"
        '                        SQLStr = "iRet=" & "update" & "_" & LCase(print_table.table_name) & "_" & Format(CStr(j), "00") & "(" & program.tables_struct_update_argv.Item(i + 1) & ")" & ";"
        '                        Print #1, Chr(9) & SQLStr
        '                        Print #1, Chr(9) & "if(iRet!=0)"
        '                        Print #1, Chr(9) & "{"
        '                        Print #1, Chr(9) & Chr(9) & "printf(""" & SQLStr & "Error!" & "i_rc=[%d];"", iRet);"
        '                        Print #1, Chr(9) & Chr(9) & "return iRet;"
        '                        Print #1, Chr(9) & "}"
        '                        Print #1, ""
        '                    Case "删除"
        '                        SQLStr = "iRet=" & "delete" & "_" & LCase(print_table.table_name) & "_" & Format(CStr(j), "00") & "(" & program.tables_struct_delete_argv.Item(i + 1) & ")" & ";"
        '                        Print #1, Chr(9) & SQLStr
        '                        Print #1, Chr(9) & "if(iRet!=0)"
        '                        Print #1, Chr(9) & "{"
        '                        Print #1, Chr(9) & Chr(9) & "printf(""" & SQLStr & "Error!" & "i_rc=[%d];"", iRet);"
        '                        Print #1, Chr(9) & Chr(9) & "return iRet;"
        '                        Print #1, Chr(9) & "}"
        '                        Print #1, ""
        '                    Case "插入"
        '                        SQLStr = "iRet=" & "insert" & "_" & LCase(print_table.table_name) & "(" & program.tables_struct_insert_argv.Item(i + 1) & ")" & ";"
        '                        Print #1, Chr(9) & SQLStr
        '                        Print #1, Chr(9) & "if(iRet!=0)"
        '                        Print #1, Chr(9) & "{"
        '                        Print #1, Chr(9) & Chr(9) & "printf(""" & SQLStr & "Error!" & "i_rc=[%d];"", iRet);"
        '                        Print #1, Chr(9) & Chr(9) & "return iRet;"
        '                        Print #1, Chr(9) & "}"
        '                        Print #1, ""
        '                    Case Else
        '                    End Select
        '
        '            Next j
        '            Print #1, ""
        '        Next i

        Print(1, Chr(9) & "LOG_INFO(""释放数据库连接!"");" & vbCrLf)
        Print(1, Chr(9) & "db_release() ;" & vbCrLf)
        Print(1, Chr(9) & "" & vbCrLf)

        Print(1, Chr(9) & "return 0;" & vbCrLf)
        Print(1, "}" & vbCrLf)

        FileClose(1)
    End Sub






End Module
