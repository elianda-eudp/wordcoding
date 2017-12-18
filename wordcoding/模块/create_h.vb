Module create_h
Sub create_h(ByVal program)

        Dim i As Integer
        Dim single_in_file As in_file
        Dim str As String
        Dim single_file As in_file

        Dim docpath As Object
        docpath = env_conf.win_work_path & "\"

        FileOpen (1,docpath & "src\" & program.program_english_name & ".h" , OpenMode.Output )

        Print(1, "# ifndef " & Chr(9) & UCase(program.program_english_name) & "_H" & vbCrLf)
        Print(1, "# define " & Chr(9) & UCase(program.program_english_name) & "_H" & Chr(9) & "1" & vbCrLf)
        Print(1, "" & vbCrLf)
        Print(1, "#include <stdio.h>" & vbCrLf)
        Print(1, "#include <stdlib.h>" & vbCrLf)
        Print(1, "#include <string.h>" & vbCrLf)
        Print(1, "#include <pubhead.h>" & vbCrLf)

        '表的头文件
        '        For i = 0 To program.tables.table_count - 1
        '           print_table = program.tables.get_single_table(i)
        '            Print #1, "#include """ & print_table.table_name & ".h"""
        '
        '        Next i
        Print(1, "#include """ & program.program_english_name & "_db.h""" & vbCrLf)

        'Print #1, "#include ""proc_comm.h"""
        'Print #1, "#include ""structdef.h"""
        'Print #1, "#include ""macroconst.h"""
        Print(1, "#include ""endday_db_diy.h""" & vbCrLf)


        Print(1, "" & vbCrLf)
        '参数的结构体

        Print(1, "" & vbCrLf)
        Print(1, "typedef struct sys_datetime{" & vbCrLf)
        Print(1, Chr(9) & "char" & Chr(9) & "ca_date" & "[8+1];" & vbCrLf)
        Print(1, Chr(9) & "char" & Chr(9) & "ca_time" & "[6+1];" & vbCrLf)
        Print(1, Chr(9) & "FD_PUBLICPARA_REC st_para;" & vbCrLf)
        Print(1, "} SYS_DATETIME;" & vbCrLf)

        Print(1, "" & vbCrLf)
        Print(1, "# define" & Chr(9) & "SYS_DATETIME_LEN" & Space(4) & "sizeof(SYS_DATETIME)" & vbCrLf)

        Print(1, "" & vbCrLf)



        Print(1, "" & vbCrLf)
        Print(1, "typedef struct " & LCase(program.program_english_name) & "_argv{" & vbCrLf)

        For i = 1 To program.call_method.oper_count
            Print(1, Chr(9) & LCase(program.call_method.val_type_pre(i)) & Chr(9) & program.call_method.val_name(i) & "[" & program.call_method.val_type_sux(i) & "+1];" & Chr(9) & "/*  " & program.call_method.argv_note(i) & " */" & vbCrLf)
        Next i
        Print(1, Chr(9) & "int      index[128];" & Chr(9) & Chr(9) & "/*  应用参数标识 */" & vbCrLf)
        Print(1, "} " & UCase(program.program_english_name) & "_ARGV_REC;" & vbCrLf)

        Print(1, "" & vbCrLf)
        Print(1, "# define" & Chr(9) & UCase(program.program_english_name) & "_ARGV_REC_LEN" & Space(4) & "sizeof(" & UCase(program.program_english_name) & "_ARGV_REC)" & vbCrLf)

        Print(1, "" & vbCrLf)
        REM 输入文件的结构体定义

        REM 文件头结构体

        For i = 1 To program.in_file.Count
           single_in_file = New in_file
           single_in_file = program.in_file.Item(i)
            Print(1, "" & vbCrLf)
            Print(1, "typedef struct " & LCase(single_in_file.file_name) & "_format{" & vbCrLf)
            Print(1, Chr(9) & "char" & Chr(9) & "file_type" & "[20+1];" & Chr(9) & "/*  " & single_in_file.file_formats.file_type_com & " */" & vbCrLf)
            Print(1, Chr(9) & "char" & Chr(9) & "file_decollator" & "[20+1];" & Chr(9) & "/*  " & single_in_file.file_formats.file_decollator_com & " */" & vbCrLf)
            Print(1, Chr(9) & "char" & Chr(9) & "have_or_not_file_head" & "[20+1];" & Chr(9) & "/*  " & single_in_file.file_formats.have_or_not_file_head_com & " */" & vbCrLf)
            Print(1, Chr(9) & "char" & Chr(9) & "file_note" & "[20+1];" & Chr(9) & "/*  " & single_in_file.file_formats.file_note_com & " */" & vbCrLf)
            Print(1, "} " & UCase(single_in_file.file_name) & "_FORMAT_REC;" & vbCrLf)

            Print(1, "" & vbCrLf)
            Print(1, "# define" & Chr(9) & UCase(single_in_file.file_name) & "_FORMAT_REC_LEN" & Space(4) & "sizeof(" & UCase(single_in_file.file_name) & "_FORMAT_REC)" & vbCrLf)

            Print(1, "" & vbCrLf)

            REM 文件头
            Print(1, "" & vbCrLf)
            Print(1, "typedef struct " & LCase(single_in_file.file_name) & "_head{" & vbCrLf)
            For j = 1 To single_in_file.file_heads.str_name.Count
                str = single_in_file.file_heads.str_type.Item(j)
                If InStr(str, "double") > 0 Or InStr(str, "int") > 0 Then
                    Print(1, Chr(9) & "char" & Chr(9) & "c_" & LCase(single_in_file.file_heads.str_name(j)) & "[40+1];" & Chr(9) & "/*  " & single_in_file.file_heads.str_note(j) & " */" & vbCrLf)
                End If
                Print(1, Chr(9) & LCase(single_in_file.file_heads.val_type_pre.Item(j)) & Chr(9) & LCase(single_in_file.file_heads.str_name(j)) & single_in_file.file_heads.val_type_sux(j) & ";" & Chr(9) & "/*  " & single_in_file.file_heads.str_note(j) & " */" & vbCrLf)
            Next j
            Print(1, "} " & UCase(single_in_file.file_name) & "_HEAD_REC;" & vbCrLf)

            Print(1, "" & vbCrLf)
            Print(1, "# define" & Chr(9) & UCase(single_in_file.file_name) & "_HEAD_REC_LEN" & Space(4) & "sizeof(" & UCase(single_in_file.file_name) & "_HEAD_REC)" & vbCrLf)

            Print(1, "" & vbCrLf)

            Print(1, "" & vbCrLf)
            Print(1, "typedef struct " & LCase(single_in_file.file_name) & "_body{" & vbCrLf)
            For j = 1 To single_in_file.file_bodys.str_name.Count
                str = single_in_file.file_bodys.str_type.Item(j)
                If InStr(str, "double") > 0 Or InStr(str, "int") > 0 Then
                    Print(1, Chr(9) & "char" & Chr(9) & "c_" & LCase(single_in_file.file_bodys.str_name(j)) & "[40+1];" & Chr(9) & "/*  " & single_in_file.file_bodys.str_note(j) & " */" & vbCrLf)
                End If
                Print(1, Chr(9) & LCase(single_in_file.file_bodys.val_type_pre.Item(j)) & Chr(9) & LCase(single_in_file.file_bodys.str_name.Item(j)) & single_in_file.file_bodys.val_type_sux.Item(j) & ";" & Chr(9) & "/*  " & single_in_file.file_bodys.str_note.Item(j) & " */" & vbCrLf)

            Next j
            Print(1, "} " & UCase(single_in_file.file_name) & "_BODY_REC;" & vbCrLf)

            Print(1, "" & vbCrLf)
            Print(1, "# define" & Chr(9) & UCase(single_in_file.file_name) & "_BODY_REC_LEN" & Space(4) & "sizeof(" & UCase(single_in_file.file_name) & "_BODY_REC)" & vbCrLf)

            Print(1, "" & vbCrLf)

            REM 批量数据集
            Print(1, "" & vbCrLf)
            Print(1, "typedef struct " & LCase(single_in_file.file_name) & "_body_set{" & vbCrLf)
            For j = 1 To single_in_file.file_bodys.str_name.Count
                str = single_in_file.file_bodys.str_type.Item(j)
                If InStr(str, "double") > 0 Or InStr(str, "int") > 0 Then
                    Print(1, Chr(9) & "char" & Chr(9) & "c_" & LCase(single_in_file.file_bodys.str_name(j)) & "[40+1];" & Chr(9) & "/*  " & single_in_file.file_bodys.str_note(j) & " */" & vbCrLf)
                End If
                Print(1, Chr(9) & LCase(single_in_file.file_bodys.val_type_pre.Item(j)) & Chr(9) & LCase(single_in_file.file_bodys.str_name.Item(j)) & "[5000]" & single_in_file.file_bodys.val_type_sux.Item(j) & ";" & Chr(9) & "/*  " & single_in_file.file_bodys.str_note.Item(j) & " */" & vbCrLf)

            Next j
            Print(1, "} " & UCase(single_in_file.file_name) & "_BODY_SET_REC;" & vbCrLf)

            Print(1, "" & vbCrLf)
            Print(1, "# define" & Chr(9) & UCase(single_in_file.file_name) & "_BODY_SET_REC_LEN" & Space(4) & "sizeof(" & UCase(single_in_file.file_name) & "_BODY_SET_REC)" & vbCrLf)

            Print(1, "" & vbCrLf)
        Next i


        For i = 1 To program.out_file.Count
           single_in_file = New in_file
           single_in_file = program.out_file.Item(i)
            Print(1, "" & vbCrLf)
            Print(1, "typedef struct " & LCase(single_in_file.file_name) & "_format{" & vbCrLf)
            Print(1, Chr(9) & "char" & Chr(9) & "file_type" & "[20+1];" & Chr(9) & "/*  " & single_in_file.file_formats.file_type_com & " */" & vbCrLf)
            Print(1, Chr(9) & "char" & Chr(9) & "file_decollator" & "[20+1];" & Chr(9) & "/*  " & single_in_file.file_formats.file_decollator_com & " */" & vbCrLf)
            Print(1, Chr(9) & "char" & Chr(9) & "have_or_not_file_head" & "[20+1];" & Chr(9) & "/*  " & single_in_file.file_formats.have_or_not_file_head_com & " */" & vbCrLf)
            Print(1, Chr(9) & "char" & Chr(9) & "file_note" & "[20+1];" & Chr(9) & "/*  " & single_in_file.file_formats.file_note_com & " */" & vbCrLf)
            Print(1, "} " & UCase(single_in_file.file_name) & "_FORMAT_REC;" & vbCrLf)

            Print(1, "" & vbCrLf)
            Print(1, "# define" & Chr(9) & UCase(single_in_file.file_name) & "_FORMAT_REC_LEN" & Space(4) & "sizeof(" & UCase(single_in_file.file_name) & "_FORMAT_REC)" & vbCrLf)

            Print(1, "" & vbCrLf)

            REM 文件头
            Print(1, "" & vbCrLf)
            Print(1, "typedef struct " & LCase(single_in_file.file_name) & "_head{" & vbCrLf)
            For j = 1 To single_in_file.file_heads.str_name.Count
                Print(1, Chr(9) & LCase(single_in_file.file_heads.val_type_pre.Item(j)) & Chr(9) & single_in_file.file_heads.str_name(j) & single_in_file.file_heads.val_type_sux(j) & ";" & Chr(9) & "/*  " & single_in_file.file_heads.str_note(j) & " */" & vbCrLf)

            Next j
            Print(1, "} " & UCase(single_in_file.file_name) & "_HEAD_REC;" & vbCrLf)

            Print(1, "" & vbCrLf)
            Print(1, "# define" & Chr(9) & UCase(single_in_file.file_name) & "_HEAD_REC_LEN" & Space(4) & "sizeof(" & UCase(single_in_file.file_name) & "_HEAD_REC)" & vbCrLf)

            Print(1, "" & vbCrLf)

            Print(1, "" & vbCrLf)
            Print(1, "typedef struct " & LCase(single_in_file.file_name) & "_body{" & vbCrLf)
            For j = 1 To single_in_file.file_bodys.str_name.Count
                Print(1, Chr(9) & LCase(single_in_file.file_bodys.val_type_pre.Item(j)) & Chr(9) & single_in_file.file_bodys.str_name.Item(j) & single_in_file.file_bodys.val_type_sux.Item(j) & ";" & Chr(9) & "/*  " & single_in_file.file_bodys.str_note.Item(j) & " */" & vbCrLf)

            Next j
            Print(1, "} " & UCase(single_in_file.file_name) & "_BODY_REC;" & vbCrLf)

            Print(1, "" & vbCrLf)
            Print(1, "# define" & Chr(9) & UCase(single_in_file.file_name) & "_BODY_REC_LEN" & Space(4) & "sizeof(" & UCase(single_in_file.file_name) & "_BODY_REC)" & vbCrLf)

            Print(1, "" & vbCrLf)

            REM 批量数据集
            Print(1, "" & vbCrLf)
            Print(1, "typedef struct " & LCase(single_in_file.file_name) & "_body_set{" & vbCrLf)
            For j = 1 To single_in_file.file_bodys.str_name.Count
                Print(1, Chr(9) & LCase(single_in_file.file_bodys.val_type_pre.Item(j)) & Chr(9) & single_in_file.file_bodys.str_name.Item(j) & "[5000]" & single_in_file.file_bodys.val_type_sux.Item(j) & ";" & Chr(9) & "/*  " & single_in_file.file_bodys.str_note.Item(j) & " */" & vbCrLf)

            Next j
            Print(1, "} " & UCase(single_in_file.file_name) & "_BODY_SET_REC;" & vbCrLf)

            Print(1, "" & vbCrLf)
            Print(1, "# define" & Chr(9) & UCase(single_in_file.file_name) & "_BODY_SET_REC_LEN" & Space(4) & "sizeof(" & UCase(single_in_file.file_name) & "_BODY_SET_REC)" & vbCrLf)

            Print(1, "" & vbCrLf)


        Next i

        REM 输入结构体
        Print(1, "" & vbCrLf)
        Print(1, "typedef struct " & LCase(program.program_english_name) & "_handle_argv_in{" & vbCrLf)
        Print(1, Chr(9) & program.call_method_struct_val & ";" & vbCrLf)
        Print(1, Chr(9) & program.call_method_ptr_struct_val & ";" & vbCrLf)
        Print(1, Chr(9) & program.call_method_struct_val_count & ";" & vbCrLf)
        Print(1, Chr(9) & "SYS_DATETIME sys_date; " & Space(25) & "/* 公共参数表信息 */" & vbCrLf)

        For i = 1 To program.tables_struct_val.Count
            Print(1, Chr(9) & program.tables_struct_val.Item(i) & ";" & vbCrLf)
            Print(1, Chr(9) & program.tables_struct_val.Item(i) & "1;" & vbCrLf)
            Print(1, Chr(9) & program.tables_struct_val.Item(i) & "2;" & vbCrLf)
            Print(1, Chr(9) & program.tables_ptr_struct_val.Item(i) & ";" & vbCrLf)
            Print(1, Chr(9) & program.tables_struct_val_count.Item(i) & ";" & vbCrLf)
            Print(1, "" & vbCrLf)

        Next i
        
        
        Rem 文件变量定义
        For i = 1 To program.in_file.Count

            single_file = New in_file
            single_file = program.in_file.Item(i)
            Print(1, Chr(9) & UCase(single_file.file_name) & "_FORMAT_REC" & Space(1) & single_file.file_formats_val & ";" & vbCrLf)
            Print(1, Chr(9) & UCase(single_file.file_name) & "_HEAD_REC" & Space(1) & single_file.file_heads_val & ";" & vbCrLf)
            Print(1, Chr(9) & UCase(single_file.file_name) & "_BODY_REC *" & Space(1) & "p_" & single_file.file_bodys_val & " ;" & vbCrLf)
            Print(1, Chr(9) & "char file_name_" & i & "[1024+1];" & vbCrLf)

            Print(1, "" & vbCrLf)
        Next i

        Print(1, "} " & UCase(program.program_english_name) & "_HANDLE_ARGV_IN_REC;" & vbCrLf)
        Print(1, "" & vbCrLf)
        Print(1, "# define" & Chr(9) & UCase(program.program_english_name) & "_HANDLE_ARGV_IN_REC_LEN" & Space(4) & "sizeof(" & UCase(program.program_english_name) & "_HANDLE_ARGV_IN_REC)" & vbCrLf)


        REM 输出结构体
        Print(1, "" & vbCrLf)
        Print(1, "typedef struct " & LCase(program.program_english_name) & "_handle_argv_out{" & vbCrLf)
        Print(1, Chr(9) & program.call_method_struct_val & ";" & vbCrLf)
        Print(1, Chr(9) & program.call_method_ptr_struct_val & ";" & vbCrLf)
        Print(1, Chr(9) & program.call_method_struct_val_count & ";" & vbCrLf)

        For i = 1 To program.tables_struct_val.Count
            Print(1, Chr(9) & program.tables_struct_val.Item(i) & ";" & vbCrLf)
            Print(1, Chr(9) & program.tables_struct_val.Item(i) & "1;" & vbCrLf)
            Print(1, Chr(9) & program.tables_struct_val.Item(i) & "2;" & vbCrLf)
            Print(1, Chr(9) & program.tables_ptr_struct_val.Item(i) & ";" & vbCrLf)
            Print(1, Chr(9) & program.tables_struct_val_count.Item(i) & ";" & vbCrLf)
            Print(1, "" & vbCrLf)

        Next i
        
        
        Rem 文件变量定义
        
        Rem 输出文件
        For i = 1 To program.out_file.Count
           single_file = New in_file
           single_file = program.out_file.Item(i)
            Print(1, Chr(9) & UCase(single_file.file_name) & "_FORMAT_REC" & Space(1) & single_file.file_formats_val & ";" & vbCrLf)
            Print(1, Chr(9) & UCase(single_file.file_name) & "_HEAD_REC" & Space(1) & single_file.file_heads_val & ";" & vbCrLf)
            Print(1, Chr(9) & UCase(single_file.file_name) & "_BODY_REC *" & Space(1) & "p_" & single_file.file_bodys_val & " ;" & vbCrLf)
            Print(1, Chr(9) & "char file_name_" & i & "[1024+1];" & vbCrLf)

            Print(1, "" & vbCrLf)
        Next i
        Print(1, "} " & UCase(program.program_english_name) & "_HANDLE_ARGV_OUT_REC;" & vbCrLf)

        Print(1, "" & vbCrLf)
        Print(1, "# define" & Chr(9) & UCase(program.program_english_name) & "_HANDLE_ARGV_OUT_REC_LEN" & Space(4) & "sizeof(" & UCase(program.program_english_name) & "_HANDLE_ARGV_OUT_REC)" & vbCrLf)

        Print(1, "" & vbCrLf)


        Print(1, "" & vbCrLf)
        Print(1, "# endif" & vbCrLf)
        FileClose(1)
    End Sub


End Module
