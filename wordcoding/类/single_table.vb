Public Class single_table
    Public table_name As String
    Public oper_count As Integer
    Public order_no As New Collection
    Public oper_type As New Collection
    Public oper_str As New Collection
    Public where_str As New Collection
    Public order_str As New Collection
    Public note_str As New Collection

    Public Property init_count()
        Get
            Return oper_count
        End Get
        Set(ByVal value)
            oper_count = value
        End Set
    End Property

End Class
