Public Class Tables_Oper

    Public table_count As Integer      '数据表的数量
    Dim x As Long
    Dim i As Integer
    Dim tables_name() As single_table       '数据表的表名列表



    '    Public 单表 As 单表      '数据表的数量



    '    Public Function 单表数量() As Double
    '        table_count = 单表.Count
    '    End Function


    Public Sub set_tables_name(table_name As single_table)
        'MsgBox "输出的表数量" & table_count
        ReDim Preserve tables_name(table_count + 1)

        tables_name(table_count) = table_name
        'MsgBox tables_name(table_count).table_name
        'For i = 0 To table_count
        'MsgBox tables_name(i).table_name
        'Print #2, "类", i, tables_name(i).table_name
        'Next i
        table_count = table_count + 1

    End Sub




    Public Property Author() As Integer
        Get
            Return table_count
        End Get
        Set(ByVal value As Integer)
            table_count = value
        End Set
    End Property

    Public Property get_single_table(Index As Integer)
        Get
            get_single_table = tables_name(Index)
        End Get
        Set(value)

        End Set
    End Property


    Public Function set_tables(table_name As single_table)
        'MsgBox "输出的表数量" & table_count
        ReDim Preserve tables_name(table_count + 1)

        tables_name(table_count) = table_name
        'MsgBox tables_name(table_count).table_name
        'For i = 0 To table_count
        'MsgBox tables_name(i).table_name
        'Print #2, "类", i, tables_name(i).table_name
        'Next i
        table_count = table_count + 1
        Return 0
    End Function


End Class
