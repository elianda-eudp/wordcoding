Module save_utf_8
    '工程要引用  Microsoft ActiveX Data Objects 2.5，下面两个通用方法建议放在模块中。'
    Sub SaveUTF8(ByVal Text As String, ByVal FileName As String)
        Dim oStream As ADODB.Stream

        oStream = New ADODB.Stream
        oStream.Open
        oStream.Charset = "UTF-8"
        oStream.Type = ADODB.StreamTypeEnum.adTypeText
        oStream.WriteText(Text)
        oStream.SaveToFile(FileName, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)
        oStream.Close

    End Sub

    Function LoadUTF8(ByVal FileName As String) As String
        Dim oStream As ADODB.Stream

        oStream = New ADODB.Stream
        oStream.Open
        oStream.Charset = "UTF-8"
        oStream.LoadFromFile(FileName)

        LoadUTF8 = oStream.ReadText()

        oStream.Close
    End Function
End Module
