Module tables_oper_get
    Private KeyType As String

    Function tables_info_get(word_table As Object) As single_table
        Dim RowsCnt As Integer
        Dim OperStr As String
        Dim z As Integer
        Dim TableName(200) As String
        Dim WhereType As String
        Dim table_class As New single_table
        Dim data_tables = New Tables_Oper




        With word_table
            RowsCnt = word_table.Rows.Count
            table_class.oper_count = RowsCnt
            Print(2, RowsCnt & vbCrLf)
            For z = 2 To RowsCnt
                OperStr = Replace(.Cell(z, 2).Range.Text, Chr(13) & Chr(7), "")  '操作类型
                Print(2, OperStr & vbCrLf)
                table_class.oper_type.Add(OperStr)

                KeyType = Replace(.Cell(z, 3).Range.Text, Chr(13), ":")
                KeyType = Replace(KeyType, Chr(7), "") '操作列关键字
                Print(2, KeyType & vbCrLf)

                KeyType = Mid(KeyType, 1, Len(KeyType) - 1)
                Print(2, "操作列关键字" & KeyType & vbCrLf)
                table_class.oper_str.Add(KeyType)

                WhereType = Replace(.Cell(z, 4).Range.Text, Chr(13), ":")  'Where条件
                WhereType = Replace(WhereType, Chr(7), "")  'Where条件
                Print(2, WhereType & vbCrLf)
                WhereType = Mid(WhereType, 1, Len(WhereType) - 1)
                Print(2, "Where条件" & WhereType & vbCrLf)
                table_class.where_str.Add(WhereType)


                WhereType = Replace(.Cell(z, 5).Range.Text, Chr(13), ":")  'order条件
                WhereType = Replace(WhereType, Chr(7), "")  'order条件
                Print(2, WhereType & vbCrLf)
                WhereType = Mid(WhereType, 1, Len(WhereType) - 1)
                Print(2, "order条件" & WhereType & vbCrLf)
                table_class.order_str.Add(WhereType)

                WhereType = Replace(.Cell(z, 6).Range.Text, Chr(13), ":")  '备注
                WhereType = Replace(WhereType, Chr(7), "")  '备注
                Print(2, WhereType & vbCrLf)
                WhereType = Mid(WhereType, 1, Len(WhereType) - 1)
                Print(2, "备注" & WhereType & vbCrLf)
                table_class.note_str.Add(WhereType)
            Next z
        End With

        tables_info_get = table_class



    End Function









End Module
