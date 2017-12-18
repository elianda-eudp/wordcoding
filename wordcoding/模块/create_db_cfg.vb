Module create_db_cfg
Sub create_db_cfg(program As program_class)

    Dim oper_count As Integer
    Dim table_name As String
    Dim docpath As String
    Dim SQLStr As String
        Dim i As Integer, j As Integer
        Dim table As single_table


        docpath = env_conf.win_work_path & "\dbcfg\"

        If Dir(docpath, vbDirectory) = "" Then
            MkDir(docpath)
        End If
        FileOpen(1, docpath & program.program_english_name & "_db" & ".cfg", OpenMode.Output)
        Print(1, "# 该配置文件通过程序设计说明书中的""数据表""部分自动生成。：" & vbCrLf)
        Print(1, "# 文件格式：" & vbCrLf)
        Print(1, "# 表名, 操作名称[，条件字段][，操作字段][，排序字段]" & vbCrLf)
        Print(1, "# 操作名称：INSERT, DELETE, UPDATE, FETCH  " & vbCrLf)
        Print(1, "# INSERT:后续无条件字段、操作字段、排序字段" & vbCrLf)
        Print(1, "# DELETE:后续只有条件字段" & vbCrLf)
        Print(1, "# UPDATE:必须有条件字段、操作字段（即修改字段），无排序字段" & vbCrLf)
        Print(1, "# FETCH：必须有条件字段、操作字段（即返回字段），排序字段可选" & vbCrLf)
        Print(1, "# create table " & program.program_english_name & " inset , update or delete funcation" & vbCrLf)

        For j = 0 To program.tables.table_count - 1
           table = New single_table
           table = program.tables.get_single_table(j)
            oper_count = table.oper_count
            table_name = table.table_name


            Print(2, oper_count & vbCrLf)
            For i = 1 To table.oper_type.Count
                SQLStr = ""
                Print(2, table.oper_str(i) & vbCrLf)

                Print(2, table.where_str(i) & vbCrLf)
                Select Case table.oper_type(i)
                    Case "查询"
                        SQLStr = "FETCH" & "," & table.where_str(i) & "," & table.oper_str(i) & "," & table.order_str(i)
                    Case "更新"
                        SQLStr = "UPDATE" & "," & table.where_str(i) & "," & table.oper_str(i)
                    Case "删除"
                        SQLStr = "DELETE" & "," & table.where_str(i)
                    Case "插入"
                        SQLStr = "INSERT"
                    Case Else
                End Select
                Print(1, table_name & "," & Replace(SQLStr, Chr(13) & Chr(7) & Chr(10), "") & vbCrLf)

            Next i
        Next j
        FileClose(1)
    End Sub

End Module
