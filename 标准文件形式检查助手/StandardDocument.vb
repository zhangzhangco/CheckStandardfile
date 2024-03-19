Imports System.Text.RegularExpressions

Public Class StandardDocument
    Public Property Code As String
    Public Property Name As String
    Public Property formatedName As String
    Public Shared isDomestic As Boolean
    ' 存储所有有效的文件名前缀
    Private Shared validCnPrefixes As String() = {"GB", "AQ", "BB", "CB", "CH", "CJ", "CY", "DA", "DB", "DL", "DY", "DZ", "EJ", "FZ", "GA", "GC", "GD", "GH", "GM", "GY", "HB", "HG", "HJ", "HS", "HY", "JB", "JC", "JG", "JR", "JT", "JY", "LB", "LD", "LS", "LY", "MH", "MT", "MZ", "NB", "NY", "QB", "QC", "QJ", "QX", "RB", "SB", "SC", "SF", "SH", "SJ", "SL", "SN", "SW", "SY", "TB", "TD", "TY", "WB", "WH", "WJ", "WM", "WS", "WW", "XB", "XF", "YB", "YC", "YD", "YJ", "YS", "YY", "YZ", "ZY", "GSB"}
    Private Shared validInterPrefixes As String() = {"ISO", "ITU", "IEC"}

    ' 构造函数，接受文件名参数并解析赋值给属性
    Public Sub New(fileName As String)
        If IsPrefixValid(fileName, validCnPrefixes) Then
            isDomestic = True
        ElseIf IsPrefixValid(fileName, validInterPrefixes) Then
            isDomestic = False
        Else
            Code = String.Empty
            Name = fileName
            Return
        End If
        formatedName = StandardDocument.FormatedFileName(fileName)

        ' 使用正则表达式来分割文件名为code和name
        Dim pattern As String = "^(.*?)[　\u2003](.+)$"
        Dim regex As New Regex(pattern)
        Dim match As Match = regex.Match(formatedName)

        If match.Success Then
            ' 如果匹配成功，第一个捕获组是code，第二个是name
            Code = match.Groups(1).Value
            Name = match.Groups(2).Value
        Else
            ' 如果没有匹配成功，整个fileName被视为name，code为空
            Code = String.Empty
            Name = formatedName
        End If
    End Sub

    Public Shared Function FormatedFileName(ByVal fileName As String) As String
        ' 初始化formattedFileName为输入的fileName
        Dim formattedFileName As String = fileName

        ' 使用正则表达式从左边首次遇见的中文字符，将它前面的一个或多个半角空格替换为一个全角空格
        ' 构建匹配中文字符前的空格的正则表达式
        Dim chineseCharRange As String = ChrW(&H4E00) & "-" & ChrW(&H9FFF) ' 通用的汉字范围
        Dim spaceBeforeChinesePattern As String = "(\s+)(?=[" & chineseCharRange & "])"

        ' 创建正则表达式对象
        Dim regEx As New System.Text.RegularExpressions.Regex(spaceBeforeChinesePattern)
        ' 执行替换操作
        formattedFileName = regEx.Replace(formattedFileName, "　")

        ' 如果文件名以任一有效前缀开头，则执行替换操作
        If isDomestic Then
            ' 替换短横线为长横线
            formattedFileName = formattedFileName.Replace("-", "—")
        End If

        ' 替换中文字符间的半角空格为全角空格
        formattedFileName = ReplaceColonBetweenChinese(formattedFileName)

        Return formattedFileName
    End Function

    Public Shared Function IsPrefixValid(ByVal fileName As String, ByVal prefixes As String()) As Boolean
        For Each prefix As String In prefixes
            If fileName.StartsWith(prefix) Then
                Return True
            End If
        Next
        Return False
    End Function

    Shared Function ReplaceColonBetweenChinese(ByVal inputString As String) As String
        ' 正则表达式模式，用于匹配两个中文字之间的空格
        Dim regexPattern As String = "([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])"

        ' 创建正则表达式对象
        Dim regEx As New System.Text.RegularExpressions.Regex(regexPattern)

        ' 执行正则表达式替换
        Return regEx.Replace(inputString, "$1　$2")
    End Function
End Class
