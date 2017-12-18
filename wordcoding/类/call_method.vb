Public Class call_method

    Public oper_count As Integer
    Public argv_flag As New Collection
    Public choose_flag As New Collection
    Public val_name As New Collection
    Public val_type As New Collection
    Public val_type_pre As New Collection
    Public val_type_sux As New Collection
    Public argv_note As New Collection


    Public Property init_count()
        Get
            Return oper_count
        End Get
        Set(ByVal value)
            oper_count = value
        End Set
    End Property



    Public Function Split_val_type()
        Dim i As Integer
        i = 0
        For i = 1 To val_type.Count
            val_type_pre.Add（Left(val_type.Item(i), InStr(val_type.Item(i), "(") - 1)）
            val_type_sux.Add（Mid(val_type.Item(i), InStr(val_type.Item(i), "(") + 1, InStr(val_type.Item(i), ")") - InStr(val_type.Item(i), "(") - 1)）
        Next i
        Return 0
    End Function

End Class
