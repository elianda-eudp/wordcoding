Public Class in_file
    Public file_name As String
    Public file_declare As String
    Public file_formats As New file_format
    Public file_heads As New file_head
    Public file_bodys As New file_head

    REM 文件相关的结构体定义语句

    Public file_formats_val As String
    Public file_formats_argv As String
    Public file_heads_val As String
    Public file_heads_argv As String
    Public file_bodys_val As String
    Public file_bodys_ptr_val As String
    Public file_bodys_ptr_argv As String
    Public file_bodys_argv As String

    Public file_formats_struct_init As String
    Public file_heads_struct_init As String
    Public file_bodys_struct_init As String



    Public Function def_file_struct_val()


        file_formats_val = LCase(file_name) & "_format"
        file_formats_argv = UCase(file_name) & "_FORMAT_REC *" & Space(1) & file_formats_val

        file_heads_val = LCase(file_name) & "_head"
        file_heads_argv = UCase(file_name) & "_HEAD_REC *" & Space(1) & file_heads_val

        file_bodys_val = LCase(file_name) & "_bodys"
        file_bodys_argv = UCase(file_name) & "_BODY_REC *" & Space(1) & file_bodys_val

        file_bodys_ptr_val = LCase(file_name) & "_bodys"
        file_bodys_ptr_argv = UCase(file_name) & "_BODY_REC **" & Space(1) & file_bodys_val

        file_formats_struct_init = "memset(&" & file_formats_val & " ,0X00,sizeof(" & UCase(file_name) & "_FORMAT_REC))"
        file_heads_struct_init = "memset(&" & file_heads_val & " ,0X00,sizeof(" & UCase(file_name) & "_HEAD_REC))"
        file_bodys_struct_init = "memset(&" & file_bodys_val & " ,0X00,sizeof(" & UCase(file_name) & "_BODY_REC))"
        Return 0
    End Function

End Class
