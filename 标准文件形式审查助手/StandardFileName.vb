Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Word

Public Class StandardFileName
    ' 存储所有有效的文件名前缀
    Dim validCnPrefixes As String() = {"GB", "AQ", "BB", "CB", "CH", "CJ", "CY", "DA", "DB", "DL", "DY", "DZ", "EJ", "FZ", "GA", "GC", "GD", "GH", "GM", "GY", "HB", "HG", "HJ", "HS", "HY", "JB", "JC", "JG", "JR", "JT", "JY", "LB", "LD", "LS", "LY", "MH", "MT", "MZ", "NB", "NY", "QB", "QC", "QJ", "QX", "RB", "SB", "SC", "SF", "SH", "SJ", "SL", "SN", "SW", "SY", "TB", "TD", "TY", "WB", "WH", "WJ", "WM", "WS", "WW", "XB", "XF", "YB", "YC", "YD", "YJ", "YS", "YY", "YZ", "ZY", "GSB"}
    Dim validItPrefixes As String() = {"ISO"}

    Public Property StandardLevel As String
    Public Property Recommand As Boolean = True
    Public Property Number As String
    Public Property Year As Integer
    Public Property Name As String


    ' 构造函数，接受文件名参数并解析赋值给属性
    Public Sub New(fileName As String)
        Dim CnPattern As String = "^([A-Z0-9]+(/[TZ])?)\s(\d+(\.\d+)?)[-—]?(\d{4})?\s+(.+)$"
        Dim ISOPattern As String = "^([A-Z]+(/[A-Z]+)?)\s(\d+([-]\d+)?)[:]?(\d{4})?\s+(.+)$"

        '设置标准层级
        extractPrefix(fileName)

        'ISO标准
        If validItPrefixes.Contains(StandardLevel) Then
            '匹配ISO规则
            ' 创建正则表达式匹配对象
            Dim regex As New Regex(ISOPattern)
            ' 进行匹配
            Dim match As Match = regex.Match(fileName)
            ' 检查是否匹配成功
            If match.Success Then
                ' 提取匹配的组
                StandardLevel = match.Groups(1).Value '如ISO、ISO/CIE
                Number = match.Groups(2).Value '如1234、2341-1
                Year = match.Groups(3).Value
                Name = match.Groups(4).Value.Trim
            End If
        End If

        '是国内标准
        If validCnPrefixes.Contains(StandardLevel.Replace("/T", "")) Then
            '匹配GB/T 1.1规则
            ' 创建正则表达式匹配对象
            Dim regex As New Regex(CnPattern)
            ' 进行匹配
            Dim match As Match = regex.Match(fileName)
            ' 检查是否匹配成功
            If match.Success Then
                ' 提取匹配的组
                StandardLevel = match.Groups(1).Value '如GB/T、DY/T
                Number = match.Groups(2).Value '如1234、291.1
                Year = match.Groups(3).Value
                Name = match.Groups(4).Value.Trim
                Recommand = StandardLevel.Contains("/T")
            End If
        End If
    End Sub

    Private Sub extractPrefix(fileName As String)
        ' 定义正则表达式模式
        Dim pattern As String = "^([A-Za-z/])\s+(.+)$"

        ' 创建正则表达式匹配对象
        Dim regex As New Regex(pattern)

        ' 进行匹配
        Dim match As Match = regex.Match(fileName)

        ' 检查是否匹配成功
        If match.Success Then
            ' 提取匹配的组,设置类属性
            StandardLevel = match.Groups(1).Value
        End If
    End Sub

    Private Sub handleCn(fileName As String)

    End Sub
End Class