Module file_info_get
    Function file_info_get_format(word_table As Object) As file_format
        Dim RowsCnt As Integer
        Dim file_type As String
        Dim file_decollator As String
        Dim have_or_not_file_head As String
        Dim file_note As String
        Dim file_class As New file_format


        With word_table
            RowsCnt = word_table.Rows.Count
            Print(2, RowsCnt & vbCrLf)
            For z = 2 To RowsCnt
                file_type = Replace(.Cell(z, 1).Range.Text, chr(13) & chr(7), "")  '文件类型
                file_class.file_type = LCase(file_type)

                file_decollator = Replace(.Cell(z, 2).Range.Text, Chr(13) & Chr(7), "")
                file_class.file_decollator = file_decollator

                have_or_not_file_head = Replace(.Cell(z, 3).Range.Text, Chr(13) & Chr(7), "")  '有无文件头
                file_class.have_or_not_file_head = have_or_not_file_head


                file_note = Replace(.Cell(z, 4).Range.Text, Chr(13) & Chr(7), "") '备注
                file_class.file_note = file_note
            Next z
        End With

        file_info_get_format = file_class
    End Function



    Function file_info_get(word_table As Object) As file_head
        Dim RowsCnt As Integer

        Dim z As Integer
        Dim str_name As String
        Dim str_type As String
        Dim str_note As String

        Dim file_class As New file_head



        With word_table
            RowsCnt = word_table.Rows.Count
            Print(2, RowsCnt & vbCrLf)
            For z = 2 To RowsCnt
                str_name = Replace(.Cell(z, 1).Range.Text, chr(13) & chr(7), "")  '字段名称
                file_class.str_name.Add(LCase(str_name))

                str_type = Replace(.Cell(z, 2).Range.Text, Chr(13) & Chr(7), "") '字段类型
                file_class.str_type.Add(LCase(str_type))

                str_note = Replace(.Cell(z, 3).Range.Text, Chr(13) & Chr(7), ":")  '说明
                file_class.str_note.Add(str_note)

            Next z
        End With

        file_info_get = file_class



    End Function


End Module
