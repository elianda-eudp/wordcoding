Public Class file_format

    Public file_type As String
    Public file_decollator As String
    Public have_or_not_file_head As String
    Public file_note As String

    Public file_type_com As String
    Public file_decollator_com As String
    Public have_or_not_file_head_com As String
    Public file_note_com As String

    Public Function init()
        file_type_com = "文件类型"
        file_decollator_com = "分隔符"
        have_or_not_file_head_com = "有无文件头"
        file_note_com = "备注"
        Return 0
    End Function

End Class
