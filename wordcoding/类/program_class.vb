Public Class program_class
    Public program_name As String
    Public program_english_name As String
    Public program_describe As String
    Public call_method As call_method
    Public in_file As New Collection
    Public out_file As New Collection
    Public tables As Tables_Oper
    Public program_cust_handle As New customization_handle



    Public log_name As String
    Public log_level As String
    Public run_date As String
    Public event_depent As String
    Public file_depent As String
    Public bussiness_hangdle As String


    Public call_method_struct_val As String
    Public call_method_ptr_struct_val As String
    Public call_method_struct_val_count As String
    Public call_method_struct_val_init As String

    Public tables_struct_val As New Collection
    Public tables_ptr_struct_val As New Collection
    Public tables_struct_val_count As New Collection
    Public tables_struct_val_init As New Collection
    Public tables_struct_fetch_argv As New Collection
    Public tables_struct_update_argv As New Collection
    Public tables_struct_delete_argv As New Collection
    Public tables_struct_insert_argv As New Collection





    Public Function init_struct()

        Dim i As Integer, j As Integer
        Dim print_table As New single_table
        Dim comments_str As New comments_class
        Dim single_file As in_file

        call_method_struct_val = UCase(program_english_name) & "_ARGV_REC" & Chr(9) & LCase(program_english_name) & "_argv_rec"
        call_method_ptr_struct_val = UCase(program_english_name) & "_ARGV_REC *" & Chr(9) & "p_" & LCase(program_english_name) & "_argv_rec"
        call_method_struct_val_count = "int" & Chr(9) & "i_" & LCase(program_english_name) & "_argv_rec "
        call_method_struct_val_init = "memset(&" & LCase(program_english_name) & "_argv_rec ,0X00,sizeof(" & UCase(program_english_name) & "_ARGV_REC))"

        program_cust_handle.func_argv.Add(LCase(program_english_name) & "_argv_rec")
        program_cust_handle.func_argv_class.Add(UCase(program_english_name) & "_ARGV_REC" & " * " & LCase(program_english_name) & "_argv_rec")
        comments_str = New comments_class
        For i = 1 To call_method.argv_flag.Count
            comments_str.str_type.Add(call_method.val_type_pre.Item(i))
            comments_str.str_def.Add(call_method.val_name.Item(i) & "[" & call_method.val_type_sux.Item(i) & "+1]")
            comments_str.str_note.Add(call_method.argv_note.Item(i))
        Next i
        program_cust_handle.argv_comments.Add(comments_str)


        program_cust_handle.func_argv.Add("p_" & LCase(program_english_name) & "_argv_rec")
        program_cust_handle.func_argv_class.Add(UCase(program_english_name) & "_ARGV_REC ** " & "p_" & LCase(program_english_name) & "_argv_rec")
        comments_str = New comments_class
        comments_str.str_type.Add("二维数组")
        comments_str.str_def.Add("p_" & LCase(program_english_name) & "_argv_rec")
        comments_str.str_note.Add（"存放了" & "i_" & LCase(program_english_name) & "_argv_rec条" & LCase(program_english_name) & "_argv_rec记录的结果集,通过下标访问"）
        program_cust_handle.argv_comments.Add（comments_str）

        program_cust_handle.func_argv.Add（"i_" & LCase(program_english_name) & "_argv_rec"）
        program_cust_handle.func_argv_class.Add("int * " & "i_" & LCase(program_english_name) & "_argv_rec")
        comments_str = New comments_class
        comments_str.str_type.Add("计数器")
        comments_str.str_def.Add("i_" & LCase(program_english_name) & "_argv_rec")
        comments_str.str_note.Add("记录" & "p_" & LCase(program_english_name) & "_argv_rec" & "中记录的条数")
        program_cust_handle.argv_comments.Add(comments_str)


        For i = 0 To tables.table_count - 1
            print_table = tables.get_single_table(i)
            tables_struct_val.Add(UCase(print_table.table_name) & "_REC" & Chr(9) & LCase(print_table.table_name) & "_rec")
            program_cust_handle.func_argv.Add(LCase(print_table.table_name) & "_rec")
            program_cust_handle.func_argv_class.Add(UCase(print_table.table_name) & "_REC * " & LCase(print_table.table_name) & "_rec")
            comments_str = New comments_class
            For j = 1 To print_table.oper_str.Count
                comments_str.str_type.Add(print_table.oper_str.Item(j))
                comments_str.str_def.Add(print_table.oper_str.Item(j))
                comments_str.str_note.Add(print_table.oper_str.Item(j))
            Next j
            program_cust_handle.argv_comments.Add(comments_str)


            tables_ptr_struct_val.Add(UCase(print_table.table_name) & "_REC *" & Chr(9) & "p_" & LCase(print_table.table_name) & "_rec ")
            program_cust_handle.func_argv.Add("p_" & LCase(print_table.table_name) & "_rec")
            program_cust_handle.func_argv_class.Add(UCase(print_table.table_name) & "_REC ** " & "p_" & LCase(print_table.table_name) & "_rec")
            comments_str = New comments_class
            comments_str.str_type.Add("二维数组")
            comments_str.str_def.Add("p_" & LCase(print_table.table_name) & "_rec")
            comments_str.str_note.Add("存放了" & "i_" & LCase(print_table.table_name) & "_rec条" & LCase(print_table.table_name) & "_rec记录的结果集,通过下标访问")
            program_cust_handle.argv_comments.Add(comments_str)

            tables_struct_val_count.Add("int" & Chr(9) & "i_" & LCase(print_table.table_name) & "_rec ")
            program_cust_handle.func_argv.Add("i_" & LCase(print_table.table_name) & "_rec")
            program_cust_handle.func_argv_class.Add("int * " & "i_" & LCase(print_table.table_name) & "_rec")
            comments_str = New comments_class
            comments_str.str_type.Add("计数器")
            comments_str.str_def.Add("i_" & LCase(print_table.table_name) & "_rec")
            comments_str.str_note.Add("记录" & "p_" & "p_" & LCase(print_table.table_name) & "_rec" & "中记录的条数")
            program_cust_handle.argv_comments.Add(comments_str)

            tables_struct_val_init.Add("memset(&" & LCase(print_table.table_name) & "_rec ,0X00,sizeof(" & UCase(print_table.table_name) & "_REC))")
            tables_struct_fetch_argv.Add("&in_set." & LCase(print_table.table_name) & "_rec, " & "&in_set.p_" & LCase(program_english_name) & "_argv_rec , " & "&in_set.i_" & LCase(program_english_name) & "_argv_rec")
            tables_struct_update_argv.Add("&out_set." & LCase(print_table.table_name) & "_rec1, " & "&out_set." & LCase(print_table.table_name) & "_rec2")
            tables_struct_delete_argv.Add("&out_set." & LCase(print_table.table_name) & "_rec ")
            tables_struct_insert_argv.Add("&out_set." & LCase(print_table.table_name) & "_rec ")
        Next i


        For i = 1 To in_file.Count
            single_file = in_file.Item(i)
            comments_str = New comments_class
            program_cust_handle.func_argv.Add(single_file.file_bodys_val)
            program_cust_handle.func_argv_class.Add(single_file.file_bodys_argv)
            comments_str.str_type.Add("二维数组，存放了输入文件[" & single_file.file_name & "]的记录内容")
            comments_str.str_def.Add(single_file.def_file_struct_val)
            comments_str.str_note.Add("可以通过下标访问，结构和详细设计中输入文件体相同")
            program_cust_handle.argv_comments.Add(comments_str)

        Next i

        For i = 1 To out_file.Count
            single_file = out_file.Item(i)
            comments_str = New comments_class
            program_cust_handle.func_argv.Add(single_file.file_bodys_val)
            program_cust_handle.func_argv_class.Add(single_file.file_bodys_argv)
            comments_str.str_type.Add("二维数组，作为输出数据容器使用，存放了输出文件[" & single_file.file_name & "]的记录内容")
            comments_str.str_def.Add(single_file.def_file_struct_val)
            comments_str.str_note.Add("此处需要对内存进行动态分配:" & Chr(10) & Chr(9) & "p_" & single_file.file_bodys_val & " = (" & UCase(single_file.file_name) & "_BODY_REC *)calloc(atoi(" & single_file.file_heads_val & "->total_count)," & UCase(single_file.file_name) & "_BODY_REC_LEN" & ");")

            program_cust_handle.argv_comments.Add(comments_str)
        Next i
        Return 0
    End Function

End Class
