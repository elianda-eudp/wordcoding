Imports wordcoding

Module create_regedit_cfg

    Private table As single_table

    Sub create_regedit_cfg(programs As Collection)

        Dim i As Integer, j As Integer
        Dim program As New program_class
        Dim docpath As Object

        docpath = env_conf.win_work_path & "\"

        If Dir(docpath & "regedit\", vbDirectory) = "" Then
            MkDir(docpath & "regedit\")
        End If
        FileClose(1)
        FileOpen(1, docpath & "regedit\" & "prog" & "_regedit" & ".cfg", OpenMode.Output)
        FileClose(4)
        FileOpen(4, docpath & "regedit\" & "tables" & "_regedit" & ".cfg", OpenMode.Output)

        'MsgBox(programs.Count)
        For i = 1 To programs.Count
            program = programs.Item(i)
            Print(1, program.program_english_name & vbCrLf)
            For j = 0 To program.tables.table_count - 1
                table = New single_table
                table = program.tables.get_single_table(j)
                Print(4, table.table_name & vbCrLf)
            Next j
        Next i


        FileClose(1)
        FileClose(4)

    End Sub


End Module
