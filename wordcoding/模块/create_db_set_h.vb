Module create_db_set_h
Sub create_db_set_h(ByVal program)

        Dim i As Integer
        Dim single_in_file As in_file
        Dim str As String

        Dim docpath As Object
        docpath = env_conf.win_work_path & "\"
        If Dir(docpath & "src\", vbDirectory) = "" Then
            MkDir(docpath & "src\")
        End If
        FileOpen(1,docpath & "src\" & program.program_english_name & "_set.h" , OpenMode.Output )

        Print(1, "# ifndef " & Chr(9) & UCase(program.program_english_name) & "_SET_H" & vbCrLf)
        Print(1, "# define " & Chr(9) & UCase(program.program_english_name) & "_SET_H" & Chr(9) & "1" & vbCrLf)
        Print(1, "")

        Print(1, "#include """ & program.program_english_name & "_db.h""" & vbCrLf)
        Print(1, "#include ""endday_db_diy.h""" & vbCrLf)



        Print(1, "" & vbCrLf)


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
            Print(1, "} " & UCase(single_in_file.file_name) & "_FORMAT_REC;" & vbCrLf）

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






        Print(1, "" & vbCrLf)
        Print(1, "# endif" & vbCrLf)
        FileClose(1)
    End Sub



End Module
