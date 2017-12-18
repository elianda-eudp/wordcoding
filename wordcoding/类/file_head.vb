Public Class file_head
    Public str_name As New Collection
    Public str_type As New Collection
    Public str_note As New Collection

    Public val_type_pre As New Collection
    Public val_type_sux As New Collection

    Public str_name_com As String
    Public str_type_com As String
    Public str_note_com As String
    Dim str As String

    Public Function init()
        str_name_com = "字段名称"
        str_type_com = "字段类型"
        str_note_com = "说明"
        Return 0
    End Function

    Public Function split_str_type()
        Dim i As Integer
        i = 0
        For i = 1 To str_type.Count
            str = str_type.Item(i)
            If InStr(str, "double") > 0 Then
                val_type_pre.Add("double")
                val_type_sux.Add("")
            ElseIf InStr(str, "int") > 0 Then
                val_type_pre.Add("int")
                val_type_sux.Add("")
            Else
                val_type_pre.Add(Left(str_type.Item(i), InStr(str_type.Item(i), "(") - 1)）
                val_type_sux.Add（"[" & Mid(str_type.Item(i), InStr(str_type.Item(i), "(") + 1, InStr(str_type.Item(i), ")") - InStr(str_type.Item(i), "(") - 1) & "+1]"）
            End If

        Next i
        Return 0
    End Function

End Class
