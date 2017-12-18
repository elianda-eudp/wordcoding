Module create_file_c
Sub create_file_c(ByVal value)

    Dim i As Integer
    Dim j As Integer
    Dim program As New program_class
    Dim single_file As New in_file
    Dim file_str As String
    Dim str As String



        Dim docpath As Object
        docpath = env_conf.win_work_path & "\"
        program = value
        FileClose(5)
        FileOpen (5,docpath & "src\" & program.program_english_name & "_file" & ".c" , OpenMode.Output)

        Print(5, "#include """ & program.program_english_name & ".h""" & vbCrLf)
        Print(5, "" & vbCrLf)

        For i = 1 To program.in_file.Count
           single_file = program.in_file.Item(i)
            Print(5, "int " & single_file.file_name & "_get( char * pfile ," & single_file.file_formats_argv & " , " & single_file.file_heads_argv & " , " & single_file.file_bodys_ptr_argv & ")" & vbCrLf)
            Print(5, "{" & vbCrLf)
            Print(5, Chr(9) & "int iRet=0;" & vbCrLf)
            Print(5, Chr(9) & "int i=0;" & vbCrLf)
            Print(5, Chr(9) & "char c_str_buf[1024];" & vbCrLf)
            Print(5, Chr(9) & "FILE * p_in_file_" & i & " = NULL;" & vbCrLf)
            'Print #5, Chr(9) & single_file.file_formats_argv & " = NULL ;"
            Print(5, Chr(9) & UCase(single_file.file_name) & "_BODY_REC " & Space(1) & single_file.file_bodys_val & "_1;" & vbCrLf)
            Print(5, Chr(9) & UCase(single_file.file_name) & "_BODY_REC *" & Space(1) & "p_" & single_file.file_bodys_val & ";" & vbCrLf)
            Print(5, "" & vbCrLf)



            Print(5, Chr(9) & "memset(c_str_buf, 0x00, sizeof(c_str_buf));" & vbCrLf)
            Print(5, "" & vbCrLf)
            Print(5, Chr(9) & "LOG_INFO(""文件[%s]打开开始!"",pfile);" & vbCrLf)

            Print(5, Chr(9) & "//输入文件打开" & vbCrLf)
            Print(5, Chr(9) & "if ( ( p_in_file_" & i & " = fopen(pfile , ""r"" ) ) == NULL )" & vbCrLf)
            Print(5, Chr(9) & "{" & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "if(p_in_file_" & i & "!=NULL) fclose( " & "p_in_file_" & i & ");" & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "LOG_ERR(""文件[%s]打开失败!"",pfile);" & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "return (-1);" & vbCrLf)
            Print(5, Chr(9) & "}" & vbCrLf)
            Print(5, "" & vbCrLf)

            Print(5, Chr(9) & "LOG_INFO(""获取文件[%s]第一行记录!"",pfile);" & vbCrLf)

            Print(5, Chr(9) & "if(fgets( c_str_buf, sizeof( c_str_buf ) - 1, p_in_file_" & i & ") == NULL )" & vbCrLf)
            Print(5, Chr(9) & "{" & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "fclose(p_in_file_" & i & ");" & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "LOG_ERR(""文件[%s]获取文件头失败!"",pfile);" & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "return -1;" & vbCrLf)
            Print(5, Chr(9) & "}" & vbCrLf)
            Print(5, Chr(9) & "LOG_INFO(""获取文件[%s]文件的头!"",pfile);" & vbCrLf)
            Print(5, "" & vbCrLf)
            Print(5, Chr(9) & "//获取文件的头" & vbCrLf)
            For j = 1 To single_file.file_heads.str_name.Count
                str = single_file.file_heads.str_type.Item(j)
                If InStr(str, "int") > 0 Then
                    Print(5, Chr(9) & "str_get_str_sep_value(c_str_buf ," & single_file.file_heads_val & "->c_" & single_file.file_heads.str_name.Item(j) & ", """ & single_file.file_formats.file_decollator & """, " & j & " );" & vbCrLf)
                    Print(5, Chr(9) & single_file.file_heads_val & "->" & single_file.file_heads.str_name.Item(j) & "= atoi(" & single_file.file_heads_val & "->c_" & single_file.file_heads.str_name.Item(j) & ");" & vbCrLf)
                ElseIf InStr(str, "double") > 0 Then
                    Print(5, Chr(9) & "str_get_str_sep_value(c_str_buf ," & single_file.file_heads_val & "->c_" & single_file.file_heads.str_name.Item(j) & ", """ & single_file.file_formats.file_decollator & """, " & j & " );" & vbCrLf)
                    Print(5, Chr(9) & single_file.file_heads_val & "->" & single_file.file_heads.str_name.Item(j) & "= atof(" & single_file.file_heads_val & "->c_" & single_file.file_heads.str_name.Item(j) & ");" & vbCrLf)
                Else
                    Print(5, Chr(9) & "str_get_str_sep_value(c_str_buf ," & single_file.file_heads_val & "->" & single_file.file_heads.str_name.Item(j) & ", """ & single_file.file_formats.file_decollator & """, " & j & " );" & vbCrLf)
                End If
            Next j

            Print(5, Chr(9) & "LOG_INFO(""获取文件[%s]文件的头结束!"",pfile);" & vbCrLf)
            Print(5, "" & vbCrLf)

            Print(5, Chr(9) & "p_" & single_file.file_bodys_val & " = (" & UCase(single_file.file_name) & "_BODY_REC *)calloc(atoi(" & single_file.file_heads_val & "->total_count)," & UCase(single_file.file_name) & "_BODY_REC_LEN" & ");" & vbCrLf)


            Print(5, Chr(9) & "//获取文件的体" & vbCrLf)
            Print(5, Chr(9) & "for(i=0;i<atoi(" & single_file.file_heads_val & "->total_count);i++)" & vbCrLf)
            Print(5, Chr(9) & "{" & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "memset(c_str_buf, 0x00, sizeof(c_str_buf));" & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "memset(&" & single_file.file_bodys_val & "_1, 0x00," & UCase(single_file.file_name) & "_BODY_REC_LEN);" & vbCrLf)

            REM 结构体需初始化
            Print(5, Chr(9) & Chr(9) & "if ( fscanf( p_in_file_" & i & ", ""%[^\n]%*c"", c_str_buf ) != 1 ) continue;" & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "LOG_INFO(""获取文件[%s]第[%d]行为[%s]!"",pfile,i,c_str_buf);" & vbCrLf)
            For j = 1 To single_file.file_bodys.str_name.Count
                str = single_file.file_bodys.str_type.Item(j)
                If InStr(str, "int") > 0 Then
                    Print(5, Chr(9) & Chr(9) & "str_get_str_sep_value(c_str_buf ," & single_file.file_bodys_val & "_1.c_" & single_file.file_bodys.str_name.Item(j) & ", """ & single_file.file_formats.file_decollator & """, " & j & " );" & vbCrLf)
                    Print(5, Chr(9) & Chr(9) & single_file.file_bodys_val & "_1." & single_file.file_bodys.str_name.Item(j) & "= atoi(" & single_file.file_bodys_val & "_1.c_" & single_file.file_bodys.str_name.Item(j) & ");" & vbCrLf)
                ElseIf InStr(str, "double") > 0 Then
                    Print(5, Chr(9) & Chr(9) & "str_get_str_sep_value(c_str_buf ," & single_file.file_bodys_val & "_1.c_" & single_file.file_bodys.str_name.Item(j) & ", """ & single_file.file_formats.file_decollator & """, " & j & " );" & vbCrLf)
                    Print(5, Chr(9) & Chr(9) & single_file.file_bodys_val & "_1." & single_file.file_bodys.str_name.Item(j) & "= atof(" & single_file.file_bodys_val & "_1.c_" & single_file.file_bodys.str_name.Item(j) & ");" & vbCrLf)
                Else
                    Print(5, Chr(9) & Chr(9) & "str_get_str_sep_value(c_str_buf ," & single_file.file_bodys_val & "_1." & single_file.file_bodys.str_name.Item(j) & ", """ & single_file.file_formats.file_decollator & """, " & j & " );" & vbCrLf)
                End If
                Print(5, Chr(9) & Chr(9) & "LOG_INFO(""[%p]!""," & single_file.file_bodys_val & "_1." & single_file.file_bodys.str_name.Item(j) & ");" & vbCrLf)

            Next j

            Print(5, Chr(9) & Chr(9) & "p_" & single_file.file_bodys_val & "[i] = " & single_file.file_bodys_val & "_1;" & vbCrLf)
            Print(5, Chr(9) & "}" & vbCrLf)
            Print(5, "" & vbCrLf)
            Print(5, Chr(9) & "*" & single_file.file_bodys_val & "=" & "p_" & single_file.file_bodys_val & ";" & vbCrLf)

            Print(5, Chr(9) & "return 0;" & vbCrLf)
            Print(5, "}" & vbCrLf)
            Print(5, "" & vbCrLf)
            Print(5, "" & vbCrLf)

        Next i
        
        Rem 输出文件的输出
        
        For i = 1 To program.out_file.Count
           single_file = program.out_file.Item(i)
            Print(5, "int " & single_file.file_name & "_put( char * pfile ," & single_file.file_formats_argv & " , " & single_file.file_heads_argv & " , " & single_file.file_bodys_argv & ")" & vbCrLf)
            Print(5, "{" & vbCrLf)
            Print(5, Chr(9) & "int iRet=0;" & vbCrLf)
            Print(5, Chr(9) & "int i=0;" & vbCrLf)
            Print(5, Chr(9) & "char c_str_buf[1024];" & vbCrLf)
            Print(5, Chr(9) & "FILE * p_out_file_" & i & " = NULL;" & vbCrLf)
            'Print #5, Chr(9) & single_file.file_formats_argv & " = NULL ;"
            'Print #5, Chr(9) & single_file.file_bodys_argv & " = NULL ;"
            Print(5, "" & vbCrLf)

            Print(5, Chr(9) & "memset(c_str_buf, 0x00, sizeof(c_str_buf));" & vbCrLf)
            Print(5, "" & vbCrLf)

            Print(5, Chr(9) & "//输出文件写入" & vbCrLf)
            Print(5, Chr(9) & "if ( ( p_out_file_" & i & " = fopen(pfile , ""w"" ) ) == NULL )" & vbCrLf)
            Print(5, Chr(9) & "{" & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "if(p_out_file_" & i & "!=NULL) fclose( " & "p_out_file_" & i & ");" & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "LOG_ERR(""文件[%s]打开失败!"",pfile);" & vbCrLf）
            Print(5, Chr(9) & Chr(9) & "return (-1);" & vbCrLf)
            Print(5, Chr(9) & "}" & vbCrLf)
            Print(5, "" & vbCrLf)

            Print(5, "" & vbCrLf)

            Print(5, Chr(9) & "//写文件头" & vbCrLf)
            Print(5, "" & vbCrLf)
            Print（5, Chr(9) & "//获取文件的头" & vbCrLf）
            file_str = ""
            file_str = Chr(9) & "sprintf(c_str_buf ,"""
            For j = 1 To single_file.file_heads.str_name.Count
                str = single_file.file_heads.str_type.Item(j)
                If InStr(str, "double") > 0 Then
                    file_str = file_str & "%.2f" & single_file.file_formats.file_decollator
                ElseIf InStr(str, "int") > 0 Then
                    file_str = file_str & "%d" & single_file.file_formats.file_decollator
                Else
                    file_str = file_str & "%s" & single_file.file_formats.file_decollator
                End If
            Next j
            file_str = file_str & "\n""," & Chr(10) & Chr(9) & Chr(9)
            For j = 1 To single_file.file_heads.str_name.Count
                file_str = file_str & single_file.file_heads_val & "->" & single_file.file_heads.str_name.Item(j)
                If j <> single_file.file_heads.str_name.Count Then
                    file_str = file_str & ","
                End If
                
                If j Mod 5 = 0 Then
                    file_str = file_str & Chr(10) & Chr(9) & Chr(9)
                End If
            Next j
            file_str = file_str & ");"
            Print(5, file_str & vbCrLf)
            Print(5, Chr(9) & "fprintf(p_out_file_" & i & ",""%s"",c_str_buf);" & vbCrLf)
            Print(5, Chr(9) & "LOG_INFO(""文件头输出!"");" & vbCrLf)



            Print(5, Chr(9) & "//获取文件的体" & vbCrLf)
            Print(5, Chr(9) & "for(i=0;i<atoi(" & single_file.file_heads_val & "->total_count);i++)" & vbCrLf)
            Print(5, Chr(9) & "{" & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "memset(c_str_buf, 0x00, sizeof(c_str_buf));" & vbCrLf)

            REM 结构体需初始化
            file_str = ""
            file_str = "sprintf(c_str_buf ,"""
            For j = 1 To single_file.file_bodys.str_name.Count
                str = single_file.file_bodys.str_type.Item(j)
                If InStr(str, "double") > 0 Then
                    file_str = file_str & "%.2f" & single_file.file_formats.file_decollator
                ElseIf InStr(str, "int") > 0 Then
                    file_str = file_str & "%d" & single_file.file_formats.file_decollator
                Else
                    file_str = file_str & "%s" & single_file.file_formats.file_decollator
                End If
            Next j
            file_str = file_str & "\n""," & Chr(10) & Chr(9) & Chr(9)
            For j = 1 To single_file.file_bodys.str_name.Count
                file_str = file_str & single_file.file_bodys_val & "[i]." & single_file.file_bodys.str_name.Item(j)
                If j <> single_file.file_bodys.str_name.Count Then
                    file_str = file_str & ","
                End If
                
                If j Mod 5 = 0 Then
                    file_str = file_str & Chr(10) & Chr(9) & Chr(9)
                End If
                
                
            Next j
            file_str = file_str & ");"
            Print(5, Chr(9) & Chr(9) & file_str & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "fprintf(p_out_file_" & i & ",""%s"",c_str_buf);" & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "LOG_INFO("" c_str_buf[%s]"",c_str_buf);" & vbCrLf)
            Print(5, Chr(9) & Chr(9) & "LOG_INFO("" [%d]"",i);" & vbCrLf)

            Print(5, Chr(9) & "}" & vbCrLf)
            Print(5, "" & vbCrLf)
            Print(5, Chr(9) & "if(p_out_file_" & i & "!=NULL) fclose( " & "p_out_file_" & i & ");" & vbCrLf)
            Print(5, Chr(9) & "return 0;" & vbCrLf)
            Print(5, "}" & vbCrLf)

        Next i


        FileClose(5)
    End Sub







End Module
