'Public Class cppvector
'    Private av() As Variant
'    Private nLast As Long
'    Private nChunk As Long

'    Public Property Get Chunk() As Long
'   Chunk = nChunk
'End Property

'    Public Property Let Chunk(NewValue As Long)
'    If NewValue = nChunk Then Exit Property
'   If NewValue < 1 Then
'      Err.Raise vbObjectError Or 1002, "CVector.Chunk", "Chunk is less than one"
'      Exit Property
'    End If
'   nChunk = NewValue
'End Property

'    Public Property Get Last() As Long
'   Last = nLast
'End Property

'    Public Property Let Last(ByVal NewLast As Long)
'    If NewLast = nLast Then Exit Property
'   If NewLast < 1 Then Exit Property
'   ReDim Preserve av(1 To NewLast) As Variant
'   nLast = NewLast
'End Property

'    Public Property Let Item(ByVal Index As Long, ByVal V As Variant)

'    If Index < 1 Then
'      Err.Raise vbObjectError Or 1000, "CVector.Let", "Index is less than one"
'      Exit Property
'    End If

'    On Error GoTo Error_Handler

'   av(Index) = V

'   If Index > nLast Then
'      nLast = Index
'   End If

'    Exit Property

'Error_Handler:

'    If Index > UBound(av) Then
'    ReDim Preserve av(1 To Index + nChunk) As Variant
'      Resume
'    End If
'   Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

'End Property

'    Public PropertyItem(ByVal Index As Long, ByVal V As Variant)

'    If Index < 1 Then
'      Err.Raise vbObjectError Or 1000, "CVector.Let", "Index is less than one"
'      Exit Property
'    End If

'    On Error GoTo Error_Handler

'   av(Index) = V

'   If Index > nLast Then
'      nLast = Index
'   End If

'    Exit Property

'Error_Handler:

'    If Index > UBound(av) Then
'    ReDim Preserve av(1 To Index + nChunk) As Variant
'      Resume
'    End If
'   Err.Raise Err.Number, Err.Source, _
'             Err.Description, Err.HelpFile, Err.HelpContext

'End Property

'    Public Property Get Item(ByVal Index As Long) As Variant

'    On Error Resume Next

'    If Index < 1 Then
'      Err.Raise vbObjectError Or 1000, _
'                "CVector.Get", "Index is less than one"
'      Exit Property
'    End If

'    If Index > nLast Then
'      Err.Raise vbObjectError Or 1001, _
'                "CVector.Get", "Index is beyond end of vector"
'      Exit Property
'    End If

'    If IsObject(av(Index)) Then
'   Item = av(Index)
'   Else
'      Item = av(Index)
'   End If

'    End Property

'    Private Sub Class_Initialize()
'        nChunk = 10
'        nLast = 1
'        ReDim av(1 To nChunk) As Variant
'End Sub


'End Class
