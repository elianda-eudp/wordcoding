﻿'Public Class FileOperation
'    Private objTS As TextStream  ' 定义TextStream对象

'    Public Function OpenFile(strFileName As String, strMode As String)
'        Dim objFSO As FileSystemObject ' 定义文件对象

'        objFSO = New FileSystemObject
'        objTS = Nothing

'        If strMode = "R" Then
'            ' 读取方式打开文件
'            objTS = objFSO.OpenTextFile(strFileName, ForReading, True)
'        End If
'        If strMode = "W" Then
'            ' 写入方式打开文件
'            objTS = objFSO.OpenTextFile(strFileName, ForWriting, True)
'        End If
'    End Function

'    Public Function CloseFile()
'        ' 关闭文件
'        objTS.Close
'    End Function

'    Public Function GetLine() As String
'        ' 按行读取文件数据
'        GetLine = objTS.ReadLine
'    End Function

'    Public Property Get AtEndOfFile() As Boolean
'        ' 判断是否已到文件末尾
'        AtEndOfFile = objTS.AtEndOfStream()
'    End Property

'    Public Function WriteLine(strData As String)
'        ' 向文件写入一条数据
'        objTS.WriteLine(strData)
'    End Function

'    Public Function SkipLines(intLines As Integer)
'        Dim i As Integer
'        ' 跳到指定数据行，如指定行超过总行数，则指定到末尾行
'        ' 一般结合文件末尾判断函数以及读取函数使用
'        For i = 1 To intLines
'            If objTS.AtEndOfStream Then
'                Exit For
'            End If
'            objTS.SkipLine
'        Next i
'    End Function

'    Public Function SkipLastLines()
'        Dim i As Integer
'        ' 跳到指定数据行，如指定行超过总行数，则指定到末尾行
'        ' 一般结合文件末尾判断函数以及读取函数使用

'        For i = 1 To objTS.Line - 1
'            objTS.SkipLine
'        Next i
'    End Function

'End Class
