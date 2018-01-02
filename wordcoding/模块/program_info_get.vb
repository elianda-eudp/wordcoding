Imports Microsoft.Office.Interop.Word

Module program_info_get
    Private k As Integer
    Private ColsCnt As Object
    Private TmpStr As Object
    Private TabNameBegin As Integer
    Private TabNameEnd As Integer

    Sub main()
        Dim RowsCnt As Integer
        Dim note As String, SQLStr As String

        Dim t
        Dim tmp

        Dim OneTable As Object
        Dim i As Integer
        Dim j As Integer
        Dim z As Integer
        Dim flag As Integer
        Dim TableName As String
        Dim TableNameBegin, TableNameEnd As Integer

        Dim docpath As String
        Dim file_name As String
        Dim ThisDocument

        Dim programs As New Collection
        Dim program As program_class
        Dim one_call_way As New call_method
        Dim get_table As New single_table
        Dim print_table As single_table
        Dim table As single_table

        Dim single_file As in_file

        Dim argv_flag As String
        Dim choose_flag As String
        Dim val_name As String
        Dim val_type As String
        Dim argv_note As String
        Dim data_tables As Tables_Oper




        docpath = env_conf.win_work_path
        env_conf.win_sys_dir = "C:\Program Files (x86)\Microsoft\coding_create"
        'file_name = Sheet1.Range("B2").Text
        'file_name = "示例详细设计.doc"
        Call win_bat_create.win_bat_create()

        If Dir(docpath, vbDirectory) = "" Then
            MkDir(docpath)
        End If

        Dim src_dir = docpath & "\src"
        If Dir(src_dir, vbDirectory) = "" Then
            MkDir(src_dir)
        End If

        Dim remote_src_dir = src_dir & "\remote_host"
        If Dir(remote_src_dir, vbDirectory) = "" Then
            MkDir(remote_src_dir)
        End If

        FileOpen(2, src_dir & "\err.txt", OpenMode.Output）

        'ThisDocument = CreateObject("word.application")
        'ThisDocument.Documents.Open(docpath & "\" & file_name)

        ThisDocument = Globals.ThisAddIn.Application
        With ThisDocument
            i = 1

            'MsgBox ThisDocument.ActiveDocument.Paragraphs.Count
            Call get_file_from_linux.get_file_from_linux()

            Do
                program = New program_class
                data_tables = New Tables_Oper
                data_tables.Author = 0
                Print（2, ThisDocument.ActiveDocument.Paragraphs(i).Range.Text）
                t = ThisDocument.ActiveDocument.Paragraphs(i)
                flag = 0
                j = 0
                'MsgBox(t)
                If t.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1 Then
                    j = i
                    Do

                        tmp = ThisDocument.ActiveDocument.Paragraphs(j).OutlineLevel
                        'MsgBox(tmp)
                        note = ""
                        If tmp = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1 And flag <> 1 Then
                            '功能名 标题1
                            program.program_name = ThisDocument.ActiveDocument.Paragraphs(j).Range.Text


                            Dim TmpStr As Object
                            TmpStr = ThisDocument.ActiveDocument.Paragraphs(j).Range.Text
                            Dim TabNameBegin As Integer
                            'MsgBox TmpStr
                            TabNameBegin = InStr(1, TmpStr, "(")
                            Dim TabNameEnd As Integer
                            TabNameEnd = InStr(1, TmpStr, ")")
                            'If TabNameBegin <> 0 And TabNameEnd <> 0 Then
                            program.program_english_name = Mid(TmpStr, TabNameBegin + 1, TabNameEnd - TabNameBegin - 1)
                            'End If
                            flag = 1
                            j = j + 1


                        ElseIf tmp = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2 And InStr(ThisDocument.ActiveDocument.Paragraphs(j).Range.Text, "功能描述") > 0 Then
                            '功能描述 标题2
                            k = j + 1
                            Do
                                If ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "正文" Then
                                    note = note & ThisDocument.ActiveDocument.Paragraphs(k).Range.Text
                                End If

                                If ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 2" Then
                                    Exit Do
                                End If
                                k = k + 1
                            Loop
                            j = k
                            '正文
                            program.program_describe = note
                        ElseIf tmp = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2 And InStr(ThisDocument.ActiveDocument.Paragraphs(j).Range.Text, "调用方法") > 0 Then
                            '调用方法 标题2
                            'MsgBox ThisDocument.ActiveDocument.Paragraphs(j).Range.Text
                            one_call_way = New call_method
                            one_call_way.init_count = 0
                            k = j + 1
                            '表格    正文   参数标识
                            Do
                                If ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "正文" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "参数标识") > 0 Then
                                    OneTable = ThisDocument.ActiveDocument.Paragraphs(k).Range.tables(1)
                                    RowsCnt = OneTable.Rows.Count
                                    ColsCnt = OneTable.Columns.Count
                                    For m = 2 To OneTable.Rows.Count
                                        argv_flag = Replace(OneTable.Cell(m, 1).Range.Text, Chr(13) & Chr(7), "")
                                        choose_flag = Replace(OneTable.Cell(m, 2).Range.Text, Chr(13) & Chr(7), "")
                                        val_name = Replace(OneTable.Cell(m, 3).Range.Text, Chr(13) & Chr(7), "")
                                        val_type = Replace(OneTable.Cell(m, 4).Range.Text, Chr(13) & Chr(7), "")
                                        argv_note = Replace(OneTable.Cell(m, 5).Range.Text, Chr(13) & Chr(7), "")
                                        one_call_way.argv_flag.Add（LCase(argv_flag)）
                                        one_call_way.choose_flag.Add（LCase(choose_flag)）
                                        one_call_way.val_name.Add（LCase(val_name)）
                                        one_call_way.val_type.Add（LCase(val_type)）
                                        one_call_way.argv_note.Add（LCase(argv_note)）
                                        one_call_way.oper_count = one_call_way.oper_count + 1

                                    Next m
                                    k = k + RowsCnt * (ColsCnt + 1) - 1
                                End If


                                If ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 2" Then
                                    Exit Do
                                End If
                                k = k + 1
                            Loop
                            j = k

                            '输入文件 标题2
                        ElseIf tmp = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2 And InStr(ThisDocument.ActiveDocument.Paragraphs(j).Range.Text, "输入文件") > 0 Then
                            '文件名称 标题3
                            k = j + 1
                            single_file = New in_file
                            Call single_file.file_formats.init()
                            Call single_file.file_heads.init()
                            Call single_file.file_bodys.init()
                            Do
                                'MsgBox "k" & k
                                If ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 3" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "文件名称") > 0 Then
                                    'single_file = New in_file
                                    'Call single_file.file_formats.init()
                                    'Call single_file.file_heads.init()
                                    'Call single_file.file_bodys.init()

                                    TmpStr = ThisDocument.ActiveDocument.Paragraphs(k).Range.Text
                                    'MsgBox TmpStr
                                    TabNameBegin = InStr(1, TmpStr, "(")
                                    TabNameEnd = InStr(1, TmpStr, ")")
                                    single_file.file_name = Mid(TmpStr, TabNameBegin + 1, TabNameEnd - TabNameBegin - 1)
                                    'MsgBox "文件名  " & single_file.file_name
                                ElseIf ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 3" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "文件说明") > 0 Then
                                    z = k + 1
                                    note = ""
                                    Do
                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "正文" Then
                                            note = note & ThisDocument.ActiveDocument.Paragraphs(z).Range.Text
                                        End If

                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "标题 3" Then
                                            k = z - 1
                                            single_file.file_declare = note
                                            Exit Do
                                        End If
                                        z = z + 1
                                    Loop
                                    'MsgBox "文件说明  " & single_file.file_declare
                                ElseIf ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 3" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "文件格式") > 0 Then
                                    z = k + 1
                                    Do
                                        '文件格式 标题3
                                        '表格   正文 文件类型
                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "正文" And InStr(ThisDocument.ActiveDocument.Paragraphs(z).Range.Text, "文件类型") > 0 Then
                                            OneTable = ThisDocument.ActiveDocument.Paragraphs(z).Range.tables(1)
                                            RowsCnt = OneTable.Rows.Count
                                            ColsCnt = OneTable.Columns.Count
                                            single_file.file_formats = file_info_get.file_info_get_format(OneTable)
                                        End If

                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "标题 3" Then
                                            k = z - 1
                                            Exit Do
                                        End If
                                        z = z + 1
                                    Loop
                                ElseIf ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 3" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "文件头") > 0 Then
                                    '文件头  标题3
                                    '表格   正文 字段名称
                                    z = k + 1
                                    Do
                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "正文" And InStr(ThisDocument.ActiveDocument.Paragraphs(z).Range.Text, "字段名称") > 0 Then
                                            OneTable = ThisDocument.ActiveDocument.Paragraphs(z).Range.tables(1)
                                            RowsCnt = OneTable.Rows.Count
                                            ColsCnt = OneTable.Columns.Count
                                            single_file.file_heads = file_info_get.file_info_get(OneTable)
                                        End If

                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "标题 3" Then
                                            k = z - 1
                                            Exit Do
                                        End If
                                        z = z + 1
                                    Loop

                                ElseIf ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 3" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "文件体") > 0 Then
                                    '文件体 标题3
                                    '表格   正文 字段名称
                                    z = k + 1
                                    Do
                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "正文" And InStr(ThisDocument.ActiveDocument.Paragraphs(z).Range.Text, "字段名称") > 0 Then
                                            OneTable = ThisDocument.ActiveDocument.Paragraphs(z).Range.tables(1)
                                            RowsCnt = OneTable.Rows.Count
                                            ColsCnt = OneTable.Columns.Count
                                            single_file.file_bodys = file_info_get.file_info_get(OneTable)
                                        End If

                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "标题 3" Or ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "标题 2" Then
                                            k = z - 1
                                            Call single_file.file_heads.split_str_type()
                                            Call single_file.file_bodys.split_str_type()
                                            Call single_file.def_file_struct_val()

                                            If single_file.file_name <> "" Then
                                                program.in_file.Add(single_file)
                                            End If
                                            Exit Do
                                        End If
                                        z = z + 1
                                    Loop
                                End If

                                If ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 2" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "输出文件") > 0 Then
                                    j = k - 1
                                    Exit Do
                                End If
                                k = k + 1

                            Loop

                            '文件说明 标题3
                            '正文
                            '文件格式 标题3
                            '表格   正文 文件类型


                            j = j + 1
                        ElseIf tmp = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2 And InStr(ThisDocument.ActiveDocument.Paragraphs(j).Range.Text, "输出文件") > 0 Then
                            '文件名称 标题3
                            k = j + 1
                            single_file = New in_file
                            Call single_file.file_formats.init()
                            Call single_file.file_heads.init()
                            Call single_file.file_bodys.init()
                            Do
                                If ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 3" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "文件名称") > 0 Then
                                    'single_file = New in_file
                                    'Call single_file.file_formats.init()
                                    'Call single_file.file_heads.init()
                                    'Call single_file.file_bodys.init()

                                    TmpStr = ThisDocument.ActiveDocument.Paragraphs(k).Range.Text
                                    'MsgBox TmpStr
                                    TabNameBegin = InStr(1, TmpStr, "(")
                                    TabNameEnd = InStr(1, TmpStr, ")")
                                    single_file.file_name = Mid(TmpStr, TabNameBegin + 1, TabNameEnd - TabNameBegin - 1)
                                ElseIf ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 3" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "文件说明") > 0 Then
                                    z = k + 1
                                    note = ""
                                    Do
                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "正文" Then
                                            note = note & ThisDocument.ActiveDocument.Paragraphs(z).Range.Text
                                        End If

                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "标题 3" Then
                                            k = z - 1
                                            single_file.file_declare = note
                                            Exit Do
                                        End If
                                        z = z + 1
                                    Loop
                                ElseIf ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 3" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "文件格式") > 0 Then
                                    z = k + 1
                                    Do
                                        '文件格式 标题3
                                        '表格   正文 文件类型
                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "正文" And InStr(ThisDocument.ActiveDocument.Paragraphs(z).Range.Text, "文件类型") > 0 Then
                                            OneTable = ThisDocument.ActiveDocument.Paragraphs(z).Range.tables(1)
                                            RowsCnt = OneTable.Rows.Count
                                            ColsCnt = OneTable.Columns.Count
                                            single_file.file_formats = file_info_get.file_info_get_format(OneTable)
                                            'MsgBox single_file.file_formats.file_decollator
                                        End If

                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "标题 3" Then
                                            k = z - 1
                                            Exit Do
                                        End If
                                        z = z + 1
                                    Loop
                                ElseIf ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 3" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "文件头") > 0 Then
                                    '文件头  标题3
                                    '表格   正文 字段名称
                                    z = k + 1
                                    Do
                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "正文" And InStr(ThisDocument.ActiveDocument.Paragraphs(z).Range.Text, "字段名称") > 0 Then
                                            OneTable = ThisDocument.ActiveDocument.Paragraphs(z).Range.tables(1)
                                            RowsCnt = OneTable.Rows.Count
                                            ColsCnt = OneTable.Columns.Count
                                            single_file.file_heads = file_info_get.file_info_get(OneTable)
                                        End If

                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "标题 3" Then
                                            k = z - 1
                                            Exit Do
                                        End If
                                        z = z + 1
                                    Loop

                                ElseIf ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 3" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "文件体") > 0 Then
                                    '文件体 标题3
                                    '表格   正文 字段名称
                                    z = k + 1
                                    Do
                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "正文" And InStr(ThisDocument.ActiveDocument.Paragraphs(z).Range.Text, "字段名称") > 0 Then
                                            OneTable = ThisDocument.ActiveDocument.Paragraphs(z).Range.tables(1)
                                            RowsCnt = OneTable.Rows.Count
                                            ColsCnt = OneTable.Columns.Count
                                            single_file.file_bodys = file_info_get.file_info_get(OneTable)
                                        End If

                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "标题 3" Or ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "标题 2" Then
                                            k = z - 1
                                            Call single_file.file_heads.split_str_type()
                                            Call single_file.file_bodys.split_str_type()
                                            Call single_file.def_file_struct_val()

                                            If single_file.file_name <> "" Then
                                                program.out_file.Add(single_file)
                                            End If
                                            Exit Do
                                        End If
                                        z = z + 1
                                    Loop
                                End If

                                If ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 2" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "数据表") > 0 Then
                                    j = k - 1
                                    Exit Do
                                End If
                                k = k + 1

                            Loop
                            j = j + 1
                        ElseIf tmp = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2 And InStr(ThisDocument.ActiveDocument.Paragraphs(j).Range.Text, "数据表") > 0 Then
                            '数据表 标题2

                            '表名 标题3  ()
                            k = j + 1

                            Do
                                If ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 3" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "Table") > 0 Then
                                    table = New single_table
                                    table.init_count = 0
                                    TmpStr = ThisDocument.ActiveDocument.Paragraphs(k).Range.Text
                                    'MsgBox TmpStr
                                    TabNameBegin = InStr(1, TmpStr, "(")
                                    TabNameEnd = InStr(1, TmpStr, ")")
                                    TableName = Mid(TmpStr, TabNameBegin + 1, TabNameEnd - TabNameBegin - 1)

                                    'MsgBox table.table_name
                                    z = k + 1
                                    Do

                                        '表格    正文   序号

                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "正文" And InStr(ThisDocument.ActiveDocument.Paragraphs(z).Range.Text, "序号") > 0 Then
                                            OneTable = ThisDocument.ActiveDocument.Paragraphs(z).Range.tables(1)
                                            RowsCnt = OneTable.Rows.Count
                                            ColsCnt = OneTable.Columns.Count
                                            table = tables_oper_get.tables_info_get(OneTable)
                                            'MsgBox table.oper_type.Item(1)
                                            'MsgBox data_tables.table_count
                                        End If

                                        If ThisDocument.ActiveDocument.Paragraphs(z).Style.NameLocal = "标题 3" Then
                                            k = z - 1
                                            Exit Do
                                        End If
                                        z = z + 1
                                    Loop

                                    'table.table_name = TableName
                                    'MsgBox table.table_name
                                    table.table_name = TableName
                                    Call data_tables.set_tables(table)
                                End If

                                'MsgBox table.table_name

                                If ThisDocument.ActiveDocument.Paragraphs(k).Style.NameLocal = "标题 2" And InStr(ThisDocument.ActiveDocument.Paragraphs(k).Range.Text, "日志") > 0 Then
                                    j = k
                                    Exit Do
                                End If

                                k = k + 1
                            Loop
                            j = j + 1
                            '日志 标题2
                        ElseIf tmp = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2 And InStr(ThisDocument.ActiveDocument.Paragraphs(j).Range.Text, "日志") > 0 Then
                            '日志输出级别
                            '正文
                            '运行时机
                            '时点依赖
                            '事件依赖
                            '文件依赖
                            '业务处理
                            j = j + 1

                        ElseIf tmp = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1 And flag = 1 Then
                            Call one_call_way.Split_val_type()
                            program.call_method = one_call_way
                            program.tables = data_tables
                            Call program.init_struct()
                            programs.Add(program)
                            i = j
                            'MsgBox("446line" & j)
                            Exit Do
                        Else
                            j = j + 1
                        End If
                        '                    n = DateAdd("s", 1, Now)
                        '                    Do Until Now > n
                        '                    DoEvents
                        '                    Loop
                        'MsgBox j & "    " & ThisDocument.ActiveDocument.Paragraphs(j).Range.Text
                        If j = 556 Then
                            'MsgBox j & "    " & ThisDocument.ActiveDocument.Paragraphs(j).Range.Text & ThisDocument.ActiveDocument.Paragraphs(j).Style
                        End If
                    Loop
                End If
                If flag = 0 Then
                    i = i + 1
                End If
                If i >= ThisDocument.ActiveDocument.Paragraphs.Count Then
                    '            Call one_call_way.split_val_type
                    '           program.call_method = one_call_way
                    '           program.tables = data_tables
                    '            Call program.init_struct
                    '            programs.Add program
                    Exit Do
                End If

                'Print(2, ThisDocument.ActiveDocument.Paragraphs(i).Range.Text)

            Loop



            '    For i = 0 To data_tables.table_count - 1
            '       print_table = New single_table
            '       print_table = programs.Item(1).tables.get_single_table(i)
            '        Print #2, i, print_table.table_name
            '        MsgBox print_table.table_name
            '        MsgBox "需操作的字符串 " & print_table.oper_str.Item(1)
            '        Call create_db_cfg.create_db_cfg(print_table)
            '    Next i
            'MsgBox("功能数" & programs.Count)
            MsgBox(programs.Count)

            For i = 1 To programs.Count
                program = programs.Item(i)
                'MsgBox("生成数据库操作配置i=" & i)
                shell_show("create_mian_c i=" & i, "温馨提示!")
                Call create_db_cfg.create_db_cfg(program)
                'MsgBox("create_h i=" & i)
                Call create_h.create_h(program)
                'MsgBox("create_db_set_h i=" & i)
                Call create_db_set_h.create_db_set_h(program)
                'MsgBox("create_mian_c i=" & i)
                Call create_main_c.create_mian_c(program)
                'MsgBox("create_handle_c i=" & i)
                Call create_handle_c.create_handle_c(program)
                'MsgBox("create_file_c i=" & i)
                Call create_file_c.create_file_c(program)
            Next i

            Call create_regedit_cfg.create_regedit_cfg(programs)



        End With
        FileClose(2)

        On Error GoTo errorHandler
        ThisDocument.ActiveDocument.Save()
        'ThisDocument.ActiveDocument.Close()
errorHandler:
        If Err().Equals(4198) Then MsgBox("Document was not closed")

        Call send_file_to_linux.send_file_to_linux()

    End Sub
    Sub shell_show(ByVal info As String, ByVal head_info As String)
        'WsShell = CreateObject("Wscript.Shell ")
        'WsShell.Popup(info, 1, head_info)
        MsgBox(info,, head_info)

        'MessageBoxTimeout(Globals.ThisAddIn.Application.Keyboard, info, head_info, vbExclamation, 0, 3000)
    End Sub

End Module
