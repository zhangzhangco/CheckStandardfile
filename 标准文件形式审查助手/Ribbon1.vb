Imports System.Drawing
Imports System.Net.Http
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Windows.Controls
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Tools.Ribbon
Imports System.Threading.Tasks
Imports Application = Microsoft.Office.Interop.Word.Application
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports Newtonsoft.Json.Linq
Imports System.Net.Security

Public Class Ribbon1
    Private currentDoc As Word.Document
    ' 使用静态HttpClient实例以提高效率和资源复用
    Public Shared ReadOnly HttpClientInstance As New HttpClient()

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) Handles Button7.Click
        Dim aboutMessage As String = "形式审查助手" & Environment.NewLine
        aboutMessage &= "版本: 0.1.0" & Environment.NewLine
        aboutMessage &= "张鑫 WeChat：zhangxin_john" & Environment.NewLine
        aboutMessage &= "用于辅助进行标准形式审查和编制的小工具。"

        MessageBox.Show(aboutMessage, "关于", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Private Function activedoc()
        If Globals.ThisAddIn.Application.Documents.Count > 0 Then
            currentDoc = Globals.ThisAddIn.Application.ActiveDocument
        Else
            MsgBox("没有打开的文件")
            Return False
        End If
        Return True
    End Function
    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim para As Word.Paragraph
        Dim nextPara As Word.Paragraph
        Dim currentLevel As Integer
        Dim nextLevel As Integer
        Dim subLevelPara As Word.Paragraph

        If Not activedoc() Then Exit Sub
        ' 开启当前文档的修订模式
        currentDoc.TrackRevisions = True

        For Each para In currentDoc.Paragraphs
            If IsSkippedParagraph(para) Then Continue For

            currentLevel = GetLevel(para.Style.NameLocal)
            If currentLevel > 0 Then
                ' 检查标题内容是否为空
                If Trim(para.Range.Text) = vbCr Then
                    If currentLevel < 7 Then
                        para.Range.Comments.Add(Range:=para.Range, Text:="此处不能为空行（GB/T 1.1—2020的7.3）")
                    ElseIf InStr(para.Style.NameLocal, "标准文件_段") > 0 Then
                        para.Range.Delete()
                        Continue For
                    End If
                End If

                If Not para.Next Is Nothing Then
                    nextPara = para.Next
                    nextLevel = GetLevel(nextPara.Style.NameLocal)

                    ' 如果当前段落的下一个段落不是下一级别的标题
                    If nextLevel <> currentLevel + 1 Then
                        ' 检查当前级别下是否存在至少一个下级
                        If CountSubLevels(para, currentLevel) >= 1 Then
                            para.Range.Comments.Add(Range:=para.Range, Text:="该标题下存在悬置段（GB/T 1.1—2020的7.4）")
                        End If
                    End If

                    ' 如果当前级别只有一个下级
                    If CountSubLevels(para, currentLevel) = 1 Then
                        subLevelPara = FindNextSubLevel(para, currentLevel)
                        If Not subLevelPara Is Nothing Then
                            subLevelPara.Range.Comments.Add(Range:=subLevelPara.Range, Text:="冗余标题")
                        End If
                    End If
                End If
            End If
        Next para
    End Sub

    Function IsSkippedParagraph(ByVal para As Word.Paragraph) As Boolean
        ' 检查是否需要跳过的段落
        Dim skipTexts As String() = {"术语和定义", "范围", "引言", "规范性引用文件"}

        If para.Style.NameLocal = "标准文件_章标题" Then
            If skipTexts.Contains(para.Range.Text.Trim()) Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Function IsInArray(stringToBeFound As String, arr As Object) As Boolean
        ' 检查数组中是否包含特定字符串
        Dim i As Integer
        For i = LBound(arr) To UBound(arr)
            If InStr(stringToBeFound, arr(i)) Then
                IsInArray = True
                Exit Function
            End If
        Next i
        IsInArray = False
    End Function

    Function FindNextSubLevel(startPara As Paragraph, level As Integer) As Paragraph
        Dim tempPara As Paragraph
        tempPara = startPara

        Do While Not tempPara Is Nothing
            tempPara = tempPara.Next
            If GetLevel(tempPara.Style.NameLocal) = level + 1 Then
                FindNextSubLevel = tempPara
                Exit Function
            End If
        Loop

        FindNextSubLevel = Nothing
    End Function


    ' 确定段落的级别
    Function GetLevel(styleName As String) As Integer
        If InStr(styleName, "标准文件_章标题") > 0 Then
            GetLevel = 1
        ElseIf (InStr(styleName, "标准文件_一级") And InStr(styleName, "标题")) > 0 Then
            GetLevel = 2
        ElseIf (InStr(styleName, "标准文件_二级") And InStr(styleName, "标题")) > 0 Then
            GetLevel = 3
        ElseIf (InStr(styleName, "标准文件_三级") And InStr(styleName, "标题")) > 0 Then
            GetLevel = 4
        ElseIf (InStr(styleName, "标准文件_四级") And InStr(styleName, "标题")) > 0 Then
            GetLevel = 5
        ElseIf (InStr(styleName, "标准文件_五级") And InStr(styleName, "标题")) > 0 Then
            GetLevel = 6
        Else
            GetLevel = 10
        End If
    End Function

    ' 计算当前级别下有多少个下级
    Function CountSubLevels(para As Paragraph, level As Integer) As Integer
        Dim count As Integer
        Dim tempPara As Paragraph
        tempPara = para

        Do While Not tempPara.Next Is Nothing
            tempPara = tempPara.Next
            If GetLevel(tempPara.Style.NameLocal) = level + 1 Then
                count = count + 1
            ElseIf GetLevel(tempPara.Style.NameLocal) <= level Then
                Exit Do
            End If
        Loop

        CountSubLevels = count
    End Function
    ' 检查指定段落是否为下一级标题
    Function IsSubLevel(para As Paragraph, currentLevel As Integer) As Boolean
        IsSubLevel = (GetLevel(para.Style.NameLocal) = currentLevel + 1)
    End Function

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        If Not activedoc() Then Exit Sub
        ' 开启当前文档的修订模式
        currentDoc.TrackRevisions = True

        Dim para As Word.Paragraph
        Dim fileNames As New Collection
        Dim isTargetSectionStarted As Boolean
        Dim isInTargetSection As Boolean
        Dim skipFirstParagraph As Boolean
        isTargetSectionStarted = False
        isInTargetSection = False
        skipFirstParagraph = True
        Dim fullwidthSpace As String
        fullwidthSpace = ChrW(&H3000)

        ' 遍历所有段落寻找目标章节
        For Each para In currentDoc.Paragraphs
            If para.Style.NameLocal = "标准文件_章标题" Then
                If isTargetSectionStarted Then
                    ' 找到下一个章节标题，结束查找
                    Exit For
                ElseIf InStr(para.Range.Text, "规范性引用文件") > 0 Then
                    ' 找到目标章节
                    isTargetSectionStarted = True
                    isInTargetSection = True
                End If
            ElseIf isTargetSectionStarted Then
                If isInTargetSection Then
                    ' 如果需要跳过第一段，则将标志设置为False
                    If skipFirstParagraph Then
                        skipFirstParagraph = False
                    Else
                        ' 收集文件名
                        fileNames.Add(FormatFileName(para.Range.Text))
                        ' 删除文件名段落
                        para.Range.Delete()
                    End If
                ElseIf para.Style.NameLocal = "标准文件_章标题" Then
                    ' 进入新章节，退出目标章节
                    isInTargetSection = False
                End If
            End If
        Next para

        ' 插入排序后的文件名
        If isTargetSectionStarted And isInTargetSection Then
            ' 定位到目标章节标题的后面一段末尾
            Dim insertPoint As Word.Range
            Dim foundTargetSectionTitle As Boolean
            For Each para In currentDoc.Paragraphs
                If para.Style.NameLocal = "标准文件_章标题" And InStr(para.Range.Text, "规范性引用文件") > 0 Then
                    foundTargetSectionTitle = True
                ElseIf foundTargetSectionTitle Then
                    insertPoint = para.Range.Paragraphs(1).Range
                    insertPoint.Collapse(WdCollapseDirection.wdCollapseEnd)
                    Exit For
                End If
            Next para

            ' 将文件名转换为数组以便排序
            Dim arrFileNames(fileNames.Count - 1) As String
            Dim j As Long
            For j = 1 To fileNames.Count
                arrFileNames(j - 1) = fileNames.Item(j)
            Next j

            ' 使用冒泡排序对文件名数组按规则排序
            Dim sorted As Boolean
            Do
                sorted = True
                For j = 0 To UBound(arrFileNames) - 1
                    If Not IsInCustomOrder(arrFileNames(j), arrFileNames(j + 1)) Then
                        Dim temp As String
                        temp = arrFileNames(j)
                        arrFileNames(j) = arrFileNames(j + 1)
                        arrFileNames(j + 1) = temp
                        sorted = False
                    End If
                Next j
            Loop Until sorted

            ' 设置样式并逐个插入文件名文本（文本结尾无换行和回车），同时避免插入空行
            For j = 0 To UBound(arrFileNames)
                ' 检查文本是否为空再插入
                If Trim(arrFileNames(j)) <> "" AndAlso insertPoint IsNot Nothing Then
                    insertPoint.Text = arrFileNames(j)
                    '拆解文件名，得到标准编号，年份和标准名称
                    insertPoint.Style = "标准文件_段"
                    insertPoint = currentDoc.Range(insertPoint.End, insertPoint.End)
                End If
            Next j
        Else
            MsgBox("未找到'规范性引用文件'章节")
        End If
    End Sub
    '返回：
    '4位年份 — 代表查到的现行年份；
    '-1      — 代表没查到

    Public Async Function QueryStandardAsync(ByVal standardCode As String) As Task(Of String)
        Using client As New HttpClient()
            ' 替换为目标URL
            Dim url As String = "http://zxd.sacinfo.org.cn/default/com.sac.tpms.core.search.countryStdSearch.flow"

            ' 构建POST请求的内容
            Dim content As New FormUrlEncodedContent(New Dictionary(Of String, String) From {
                {"_eosFlowAction", "query"},
                {"criteria/_expr[1]/stdCode", standardCode},
                {"criteria/_expr[1]/_op", "like"}  ' 根据需要添加其他参数
            })

            ' 发送POST请求
            Dim response As HttpResponseMessage = Await client.PostAsync(url, content)

            ' 确保HTTP成功状态值
            response.EnsureSuccessStatusCode()

            ' 读取并返回响应内容
            Dim responseBody As String = Await response.Content.ReadAsStringAsync()
            Return responseBody
        End Using
    End Function
    Function FormatFileName(ByVal fileName As String) As String
        Dim formattedFileName As String
        formattedFileName = fileName

        ' 存储所有有效的文件名前缀
        Dim validPrefixes As String() = {"GB", "AQ", "BB", "CB", "CH", "CJ", "CY", "DA", "DB", "DL", "DY", "DZ", "EJ", "FZ", "GA", "GC", "GD", "GH", "GM", "GY", "HB", "HG", "HJ", "HS", "HY", "JB", "JC", "JG", "JR", "JT", "JY", "LB", "LD", "LS", "LY", "MH", "MT", "MZ", "NB", "NY", "QB", "QC", "QJ", "QX", "RB", "SB", "SC", "SF", "SH", "SJ", "SL", "SN", "SW", "SY", "TB", "TD", "TY", "WB", "WH", "WJ", "WM", "WS", "WW", "XB", "XF", "YB", "YC", "YD", "YJ", "YS", "YY", "YZ", "ZY", "GSB"}

        ' 检查文件名是否以任一有效前缀开头
        If IsPrefixValid(formattedFileName, validPrefixes) Then
            ' 替换特殊字符
            Dim yearPattern As String
            yearPattern = "(\d{4})" ' 匹配4位数字年份

            ' 替换"-"+年份为"—"+年份
            formattedFileName = formattedFileName.Replace("-", "—")

            ' 使用正则表达式替换年份后面的空格为全角空格
            Dim regEx As New System.Text.RegularExpressions.Regex(yearPattern & "\s+")
            formattedFileName = regEx.Replace(formattedFileName, "$1　")
        End If

        ' 当文件名中的两个中文字中间有“ ”时候，将“ ”替换为fullwidthSpace
        formattedFileName = ReplaceColonBetweenChinese(formattedFileName)

        Return formattedFileName
    End Function

    Function IsPrefixValid(ByVal fileName As String, ByVal prefixes As String()) As Boolean
        For Each prefix As String In prefixes
            If fileName.StartsWith(prefix) Then
                Return True
            End If
        Next
        Return False
    End Function

    Function IsInCustomOrder(ByVal str1 As String, ByVal str2 As String) As Boolean
        ' 自定义排序规则
        Dim sortOrder As String() = {"GB", "AQ", "BB", "CB", "CH", "CJ", "CY", "DA", "DB", "DL", "DY", "DZ", "EJ", "FZ", "GA", "GC", "GD", "GH", "GM", "GY", "HB", "HG", "HJ", "HS", "HY", "JB", "JC", "JG", "JR", "JT", "JY", "LB", "LD", "LS", "LY", "MH", "MT", "MZ", "NB", "NY", "QB", "QC", "QJ", "QX", "RB", "SB", "SC", "SF", "SH", "SJ", "SL", "SN", "SW", "SY", "TB", "TD", "TY", "WB", "WH", "WJ", "WM", "WS", "WW", "XB", "XF", "YB", "YC", "YD", "YJ", "YS", "YY", "YZ", "ZY", "GSB", "ISO", "IEC", "ITU", "CIE", "SMPTE", "其他"}

        Dim idx1 As Long = GetSortIndex(str1, sortOrder)
        Dim idx2 As Long = GetSortIndex(str2, sortOrder)

        Return (idx1 < idx2) Or (idx1 = idx2 AndAlso str1 < str2)
    End Function

    Function GetSortIndex(ByVal str As String, ByVal order As String()) As Long
        For i As Long = 0 To order.Length - 1
            If str.StartsWith(order(i)) Then
                Return i
            End If
        Next i
        Return order.Length - 1 ' 如果没有匹配项，返回最后一个索引
    End Function

    Function ReplaceColonBetweenChinese(ByVal inputString As String) As String
        ' 正则表达式模式，用于匹配两个中文字之间的空格
        Dim regexPattern As String = "([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])"

        ' 创建正则表达式对象
        Dim regEx As New System.Text.RegularExpressions.Regex(regexPattern)

        ' 执行正则表达式替换
        Return regEx.Replace(inputString, "$1　$2")
    End Function

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        If Not activedoc() Then Exit Sub
        Dim para As Word.Paragraph

        ' 开启当前文档的修订模式
        currentDoc.TrackRevisions = False

        '将术语内错误地被标记为一级条标题的段落变为术语条一
        UpdateParagraphStylesInDocument()

        ' 遍历每个段落
        For Each para In currentDoc.Paragraphs
            ' 检查段落是否是“标准文件_章标题”样式，且文本为“术语和定义”
            If para.Style.NameLocal = "标准文件_章标题" And Trim(para.Range.Text) = "术语和定义" & Chr(13) Then
                para = para.Next
                While Not para Is Nothing
                    Dim nextPara As Word.Paragraph = para.Next
                    If para.Style.NameLocal = "标准文件_章标题" Then Exit While

                    ' 更改“标准文件_术语条一”的段落
                    If para.Style.NameLocal = "标准文件_术语条一" Then
                        '如果这段是空,删去末尾硬回车,插入软回车
                        If String.IsNullOrWhiteSpace(para.Range.Text.Trim) Then
                            ' 获取段落末尾的范围
                            Dim paraEndRange As Word.Range = currentDoc.Range(para.Range.End - 1, para.Range.End)

                            ' 检查是否是回车符
                            If paraEndRange.Text = vbCr Then
                                ' 删除回车符
                                paraEndRange.Delete()
                                '替换成软回车
                                paraEndRange.InsertAfter(Chr(11))
                                paraEndRange.Style = currentDoc.Styles("标准文件_术语条一")
                            End If
                        End If
                        ConvertParagraphToLower(para)
                        replacewithquanjiao(para)
                        Dim tt = GetLeadingText(para)
                        If Not IsStringPresentTimes(tt, 2) Then
                            currentDoc.Comments.Add(para.Range, tt & "在文中出现少于两次，应从本章中移除。")
                        End If
                    End If
                    ' 移至下一个段落
                    para = nextPara
                End While
            End If
        Next para
    End Sub

    Private Function IsStringPresentTimes(ByVal searchString As String, ByVal times As Integer) As Boolean
        Dim count As Integer = 0
        Dim range As Word.Range = currentDoc.Content

        With range.Find
            .Text = searchString
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindStop
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            While .Execute(Forward:=True, Wrap:=Word.WdFindWrap.wdFindStop) ' 改为wdFindStop避免无限循环
                count += 1
                If count > times Then
                    Return True
                End If
                ' 更新搜索的起始位置，避免重复计数
                range.Start = range.End
                If range.Start >= currentDoc.Content.End Then
                    Exit While
                End If
                range.End = currentDoc.Content.End
            End While
        End With
        Return False ' 如果循环完成而未超过times次，返回False
    End Function
    Private Function GetLeadingText(ByVal para As Word.Paragraph) As String
        ' 定义正则表达式，移除软回车和其他可能的非打印字符
        Dim cleanText As String = System.Text.RegularExpressions.Regex.Replace(para.Range.Text, "[\v]", String.Empty)

        ' 然后，匹配非空格字符序列
        Dim pattern As String = "^[^\s　]+"
        Dim regex As New Regex(pattern)
        Dim match As Match = regex.Match(cleanText)

        If match.Success Then
            Return match.Value ' 返回匹配到的字符串
        Else
            Return String.Empty ' 如果没有匹配到，则返回空字符串
        End If
    End Function
    Private Sub replacewithquanjiao(para As Word.Paragraph)
        If para Is Nothing Then Exit Sub
        If Not String.IsNullOrWhiteSpace(para.Range.Text.Trim) Then
            ' 获取段落范围，但不包括段落末尾的特殊字符
            Dim range As Word.Range = para.Range
            range.SetRange(Start:=para.Range.Start, End:=para.Range.End)

            ' 创建正则表达式对象
            Dim pattern As String = "^(.*?)[ ]+"
            Dim replacement As String = "$1　" ' 这里的全角空格在两个引号之间
            Dim regex As New Regex(pattern)

            If Not range.Text.Contains("　") Then
                ' 正则表达式替换文本
                regex.Replace(range.Text, replacement, 1)
            End If
            ' 应用样式和格式设置到当前段落
            range.Style = "标准文件_术语条一"
            range.Font.Name = "黑体"
            range.ParagraphFormat.LeftIndent = 24
            range.ParagraphFormat.CharacterUnitFirstLineIndent = -2

            ' 注意：这种方法假设段落末尾有且仅有一个特殊字符（如 vbCr）。
            ' 如果段落结尾的处理更复杂，这种方法可能需要调整。
        End If
    End Sub
    Private Sub ConvertParagraphToLower(para As Word.Paragraph)
        If para Is Nothing Then Exit Sub
        If Not String.IsNullOrWhiteSpace(para.Range.Text.Trim) Then
            ' 获取段落范围，但不包括段落末尾的特殊字符
            Dim range As Word.Range = para.Range
            range.SetRange(Start:=para.Range.Start, End:=para.Range.End - 1)

            ' 将范围内的文本转换为小写
            Dim lowerCaseText As String = range.Text.ToLower()
            range.Text = lowerCaseText

            ' 注意：这种方法假设段落末尾有且仅有一个特殊字符（如 vbCr）。
            ' 如果段落结尾的处理更复杂，这种方法可能需要调整。
        End If
    End Sub



    Private Sub UpdateParagraphStylesInDocument()
        Dim para As Word.Paragraph
        Dim checkParagraphs As Boolean
        Dim updated As Boolean

        checkParagraphs = False
        updated = False

        ' 遍历文档中的所有段落
        For Each para In currentDoc.Paragraphs
            ' 检查段落是否是“标准文件_章标题”样式，且文本为“术语和定义”
            If para.Style.NameLocal = "标准文件_章标题" And Trim(para.Range.Text) = "术语和定义" & Chr(13) Then
                checkParagraphs = True ' 开始检查后续段落
            ElseIf para.Style.NameLocal = "标准文件_章标题" Then
                checkParagraphs = False ' 遇到下一个“标准文件_章标题”，停止检查
            ElseIf checkParagraphs And para.Style.NameLocal = "标准文件_一级条标题" Then
                ' 先将字体加粗
                ' para.Range.Font.Bold = True
                ' 然后更改段落样式
                Try
                    ' 设置段落样式
                    para.Style = currentDoc.Styles("标准文件_术语条一")
                Catch ex As Exception
                    MessageBox.Show("错误: " & ex.Message)
                End Try
                updated = True
            End If
        Next para
    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        If Not activedoc() Then Exit Sub
        Dim reftext As String

        ' 开启当前文档的修订模式
        currentDoc.TrackRevisions = True

        Dim regEx As Regex
        regEx = New Regex("(([A-Z]{2,})([_/])([A-Z])\s([0-9]{1,5}(?:\.[0-9]{1,3})?)([-—])([0-9]{4}))|(([A-Z]{2.})\s([0-9]+)(?:([-])?([0-9]))(:[0-9]{4})?)")

        reftext = extracteChapterText("规范性引用文件") & extracteChapterText("参考文献")

        ProcessParagraphs(currentDoc, regEx, reftext)
        ProcessTables(currentDoc, regEx, reftext)
    End Sub

    Function extracteChapterText(chapterTitle As String) As String
        Dim doc As Word.Document
        Dim para As Word.Paragraph
        Dim isInChapter As Boolean
        Dim extractedText As String

        ' 初始化提取文本的变量
        extractedText = ""
        isInChapter = False

        ' 设置文档对象
        doc = Globals.ThisAddIn.Application.ActiveDocument

        ' 遍历文档中的每个段落
        For Each para In doc.Paragraphs
            ' 检查段落的样式是否为"标准文件_章标题"
            If para.Style.NameLocal = "标准文件_章标题" Or para.Style.NameLocal = "标准文件_参考文献标题" Then
                ' 如果找到章节标题，检查是否与指定的章节标题匹配
                If para.Range.Text = chapterTitle & vbCr Then
                    ' 找到了指定的章节标题，开始提取文本
                    isInChapter = True
                Else
                    ' 如果不匹配指定的章节标题，停止提取文本
                    isInChapter = False
                End If
            ElseIf isInChapter Then
                ' 如果在指定章节内，将段落文本添加到提取的文本中
                extractedText = extractedText & para.Range.Text
            End If
        Next para

        ' 返回提取的文本
        extracteChapterText = extractedText
    End Function

    Private Sub ProcessParagraphs(doc As Word.Document, regEx As Regex, reftext As String)
        Dim para As Word.Paragraph
        Dim skip As Boolean
        skip = True

        For Each para In doc.Paragraphs
            If para.Style.NameLocal = "标准文件_章标题" And Trim(para.Range.Text) = "术语和定义" & Chr(13) Then
                skip = False
            ElseIf para.Style.NameLocal = "标准文件_参考文献标题" Then
                skip = True
            ElseIf skip Then
                GoTo NextParagraphDangling
            End If
            ProcessText(regEx, para.Range, reftext)
NextParagraphDangling:
        Next para
    End Sub

    Private Sub ProcessTables(doc As Word.Document, regEx As Regex, reftext As String)
        Dim tbl As Word.Table
        Dim cell As Word.Cell
        For Each tbl In doc.Tables
            For Each cell In tbl.Range.Cells
                ProcessText(regEx, cell.Range, reftext)
            Next cell
        Next tbl
    End Sub

    Private Sub ProcessText(regEx As Regex, rng As Word.Range, reftext As String)
        Dim matches As MatchCollection
        matches = regEx.Matches(rng.Text)

        Dim match As Match
        For Each match In matches
            Dim originalCode As String, modifiedCode As String
            originalCode = match.Value
            modifiedCode = originalCode

            ' 替换'_'
            If InStr(modifiedCode, "_") > 0 Then
                modifiedCode = Replace(modifiedCode, "_", "/")
            End If

            ' 替换 -
            If InStr(modifiedCode, "-") > 0 And CheckChPrefix(modifiedCode) Then
                modifiedCode = Replace(modifiedCode, "-", "—")
            End If

            ' 查找是否在规范性引用文件或参考文献里
            If Not reftext.Contains(originalCode) Or Not reftext.Contains(modifiedCode) Then
                ' 添加批注
                Dim commentRange As Word.Range
                commentRange = rng.Duplicate
                commentRange.SetRange(Start:=rng.Start + match.Index, End:=rng.Start + match.Index + match.Length)
                commentRange.Comments.Add(Range:=commentRange, Text:="在规范性引用文件和参考文献中应提及" & modifiedCode)
            End If

            If originalCode <> modifiedCode Then
                ' 定位并替换具体的匹配文本
                Dim matchRange As Word.Range
                matchRange = rng.Duplicate
                matchRange.SetRange(Start:=rng.Start + match.Index, End:=rng.Start + match.Index + match.Length)
                matchRange.Text = modifiedCode
            End If
        Next match
    End Sub

    Function CheckChPrefix(inputString As String) As Boolean
        Dim validChPrefixes As String() = {"GB", "AQ", "BB", "CB", "CH", "CJ", "CY", "DA", "DB", "DL", "DY", "DZ", "EJ", "FZ", "GA", "GC", "GD", "GH", "GM", "GY", "HB", "HG", "HJ", "HS", "HY", "JB", "JC", "JG", "JR", "JT", "JY", "LB", "LD", "LS", "LY", "MH", "MT", "MZ", "NB", "NY", "QB", "QC", "QJ", "QX", "RB", "SB", "SC", "SF", "SH", "SJ", "SL", "SN", "SW", "SY", "TB", "TD", "TY", "WB", "WH", "WJ", "WM", "WS", "WW", "XB", "XF", "YB", "YC", "YD", "YJ", "YS", "YY", "YZ", "ZY", "GSB"}

        ' 提取输入字符串中的前缀
        Dim extractedPrefix As String
        extractedPrefix = Left(inputString, InStr(inputString, " ") - 1)

        ' 检查前缀是否在validChPrefixes数组中
        Dim isPrefixValid As Boolean = validChPrefixes.Contains(extractedPrefix.ToUpper())

        Return isPrefixValid
    End Function

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click
        If Not activedoc() Then Exit Sub
        currentDoc.TrackRevisions = True ' 开启修订模式

        Dim regEx As Regex
        regEx = New Regex("\b(?<![a-zA-Z\d .:/\-—])\d{5,}(?:\.\d+)?\b(?![\-/])", RegexOptions.IgnoreCase)

        Dim paragraphs As Word.Paragraphs = currentDoc.Paragraphs
        For Each para As Word.Paragraph In paragraphs
            Dim range As Word.Range = para.Range
            Dim text As String = range.Text

            Dim matches As MatchCollection = regEx.Matches(text)
            For Each match As Match In matches
                ' 设置查找范围
                range.SetRange(match.Index + para.Range.Start, match.Index + match.Length + para.Range.Start)
                ' 替换文本
                Dim formattedNumber As String = FormatNumberWithCommas(match.Value)
                range.Text = formattedNumber
                ' 重置范围到整个段落
                range.SetRange(para.Range.Start, para.Range.End)
            Next
        Next
    End Sub


    Function FormatNumberWithCommas(num As String) As String
        Dim parts As String() = num.Split("."c)
        Dim integerPart As String = parts(0)
        Dim decimalPart As String = If(parts.Length > 1, "." + parts(1), "")

        Dim regex As Regex = New Regex("(\d)(?=(\d{3})+(?!\d))")
        integerPart = regex.Replace(integerPart, "$1,")

        Return integerPart + decimalPart
    End Function

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click
        If Not activedoc() Then Exit Sub
        ' 开启当前文档的修订模式
        currentDoc.TrackRevisions = True
        ' 保存当前的修订视图状态
        Dim originalShowRevisions As Boolean = Globals.ThisAddIn.Application.ActiveWindow.View.ShowRevisionsAndComments

        ' 设置为最终状态视图，以隐藏修订内容
        Globals.ThisAddIn.Application.ActiveWindow.View.ShowRevisionsAndComments = False

        Dim unitPairs As String
        unitPairs = "米|m|千克|kg|秒|s|安培|A|摩尔|mol|坎德拉|cd|" &
                    "牛顿|N|焦耳|J|瓦特|W|帕斯卡|Pa|伏特|V|欧姆|Ω|库仑|C|" &
                    "法拉|F|特斯拉|T|亨利|H|赫兹|Hz|勒克斯|lx|摄氏度|℃|升|l|" &
                    "克|g|毫米|mm|厘米|cm|千米|km|毫克|mg|微克|μg|吨|t|" &
                    "毫秒|ms|微秒|μs|纳秒|ns|分钟|min|小时|h|天|d|年|yr|" &
                    "华氏度|℉|巴|bar|毫米汞柱|mmHg|大气压|atm|酸碱度|pH|" &
                    "分贝|dB|弧度|rad|立体弧度|sr|流明|lm|坎德拉每平方米|cd/m2|" &
                    "电子伏特|eV|卡路里|cal|千卡路里|kcal|瓦特小时|Wh|千瓦时|kWh|" &
                    "磅力每平方英寸|psi|英里每小时|mph|帧每秒|fps|" &
                    "转每分钟|rpm|千字节|kB|兆字节|MB|吉字节|GB|太字节|TB|" &
                    "百万分之一|ppm|十亿分之一|ppb|吉帕|GPa|兆帕|Mpa|千帕|kPa|" &
                    "平方毫米|mm2|平方厘米|cm2|平方米|m2|平方千米|km2|" &
                    "平方英尺|ft2|平方码|yd2|立方毫米|mm3|立方厘米|cm3|" &
                    "立方米|m3|立方千米|km3|立方英尺|ft3|立方码|yd3|" &
                    "毫升|ml|厘升|cl|分升|dl|米每秒|m/s|米每秒平方|m/s2|" &
                    "克每立方厘米|g/cm3|千克每立方米|kg/m3|千克每升|kg/L|" &
                    "毫克每升|mg/L|微克每升|μg/L|微克每立方米|μg/m3|" &
                    "克每升|g/L|毫克每毫升|mg/ml|升每分钟|l/min|" &
                    "立方米每小时|m3/h|千瓦|kW|兆瓦|MW|吉瓦|GW|太瓦|TW|" &
                    "千伏安|kVA|兆伏安|MVA|吉伏安|GVA|公里每小时|km/h|纳米|nm|度|°|比特|bit"
        Dim unitMap As New Dictionary(Of String, String) From {
                                {"米", "m"},
                                {"千克", "kg"},
                                {"秒", "s"},
                                {"安培", "A"},
                                {"摩尔", "mol"},
                                {"坎德拉", "cd"},
                                {"牛顿", "N"},
                                {"焦耳", "J"},
                                {"瓦特", "W"},
                                {"帕斯卡", "Pa"},
                                {"伏特", "V"},
                                {"欧姆", "Ω"},
                                {"库仑", "C"},
                                {"法拉", "F"},
                                {"特斯拉", "T"},
                                {"亨利", "H"},
                                {"赫兹", "Hz"},
                                {"勒克斯", "lx"},
                                {"摄氏度", "℃"},
                                {"升", "l"},
                                {"克", "g"},
                                {"毫米", "mm"},
                                {"厘米", "cm"},
                                {"千米", "km"},
                                {"毫克", "mg"},
                                {"微克", "μg"},
                                {"吨", "t"},
                                {"毫秒", "ms"},
                                {"微秒", "μs"},
                                {"纳秒", "ns"},
                                {"分钟", "min"},
                                {"小时", "h"},
                                {"天", "d"},
                                {"年", "yr"},
                                {"华氏度", "℉"},
                                {"巴", "bar"},
                                {"毫米汞柱", "mmHg"},
                                {"大气压", "atm"},
                                {"酸碱度", "pH"},
                                {"分贝", "dB"},
                                {"弧度", "rad"},
                                {"立体弧度", "sr"},
                                {"流明", "lm"},
                                {"坎德拉每平方米", "cd/m2"},
                                {"电子伏特", "eV"},
                                {"卡路里", "cal"},
                                {"千卡路里", "kcal"},
                                {"瓦特小时", "Wh"},
                                {"千瓦时", "kWh"},
                                {"磅力每平方英寸", "psi"},
                                {"英里每小时", "mph"},
                                {"英尺每秒", "fps"},
                                {"转每分钟", "rpm"},
                                {"千字节", "kB"},
                                {"兆字节", "MB"},
                                {"吉字节", "GB"},
                                {"太字节", "TB"},
                                {"百万分之一", "ppm"},
                                {"十亿分之一", "ppb"},
                                {"吉帕", "GPa"},
                                {"兆帕", "Mpa"},
                                {"千帕", "kPa"},
                                {"平方毫米", "mm2"},
                                {"平方厘米", "cm2"},
                                {"平方米", "m2"},
                                {"平方千米", "km2"},
                                {"平方英尺", "ft2"},
                                {"平方码", "yd2"},
                                {"立方毫米", "mm3"},
                                {"立方厘米", "cm3"},
                                {"立方米", "m3"},
                                {"立方千米", "km3"},
                                {"立方英尺", "ft3"},
                                {"立方码", "yd3"},
                                {"毫升", "ml"},
                                {"厘升", "cl"},
                                {"分升", "dl"},
                                {"米每秒", "m/s"},
                                {"米每秒平方", "m/s2"},
                                {"克每立方厘米", "g/cm3"},
                                {"千克每立方米", "kg/m3"},
                                {"千克每升", "kg/L"},
                                {"毫克每升", "mg/L"},
                                {"微克每升", "μg/L"},
                                {"微克每立方米", "μg/m3"},
                                {"克每升", "g/L"},
                                {"毫克每毫升", "mg/ml"},
                                {"升每分钟", "l/min"},
                                {"立方米每小时", "m3/h"},
                                {"千瓦", "kW"},
                                {"兆瓦", "MW"},
                                {"吉瓦", "GW"},
                                {"太瓦", "TW"},
                                {"千伏安", "kVA"},
                                {"兆伏安", "MVA"},
                                {"吉伏安", "GVA"},
                                {"公里每小时", "km/h"},
                                {"纳米", "nm"},
                                {"度", "°"},
                                {"比特", "bit"}
                            }

        Dim regEx As Regex
        regEx = New Regex("([-+]?\d*\.?\d+\/?\d*)\s?(" & unitPairs & ")", RegexOptions.IgnoreCase)

        ' 遍历文档中的每个段落
        For Each para As Word.Paragraph In currentDoc.Paragraphs
            Dim matches As MatchCollection = regEx.Matches(para.Range.Text)
            Dim matchesList As List(Of Match) = matches.Cast(Of Match)().ToList()

            ' 从后向前遍历匹配项，避免索引问题
            For i As Integer = matchesList.Count - 1 To 0 Step -1
                Dim match As Match = matchesList(i)

                ' 创建一个特定于匹配项的范围
                Dim matchRange As Word.Range = currentDoc.Range(Start:=para.Range.Start + match.Index, End:=para.Range.Start + match.Index + match.Length)

                '' 如果范围内有修订，则跳过此匹配项
                'If matchRange.Revisions.Count > 0 Then Continue For

                Dim originalText As String = match.Value
                Dim newText As String
                Dim unit As String = match.Groups(2).Value

                ' 如果是中文单位，替换为英文单位符号
                If unitMap.ContainsKey(unit) Then
                    unit = unitMap(unit)
                End If

                ' 生成新文本
                newText = match.Groups(1).Value & " " & unit
                If unit = "℃" Or unit = "℉" Or unit = "°" Then
                    newText = match.Groups(1).Value & unit
                End If

                ' 检查原文本与新文本是否不同，若不同，则替换
                If originalText <> newText Then
                    matchRange.Text = newText
                End If
            Next
        Next

        ' 恢复原始的修订视图状态
        Globals.ThisAddIn.Application.ActiveWindow.View.ShowRevisionsAndComments = originalShowRevisions
    End Sub

    Private Sub Button8_Click(sender As Object, e As RibbonControlEventArgs) Handles Button8.Click
        RenewAllTableFormats()
    End Sub
    Private Sub RenewAllTableFormats()
        If Not activedoc() Then Exit Sub
        Dim tableCount As Integer = currentDoc.Tables.Count

        If tableCount = 0 Then
            MessageBox.Show("文档中没有表格。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        For i As Integer = 1 To tableCount
            Dim table As Object = currentDoc.Tables(i)

            ' 检查表格是否有框线并且行和列都大于2
            If TableHasBorders(table) AndAlso table.Rows.Count > 2 AndAlso table.Columns.Count > 2 Then
                ResetTable(table)
                SetTableFormat(table)
            End If
        Next

        MessageBox.Show("所有表格格式已更新。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Private Function TableHasBorders(ByVal table As Object) As Boolean
        Dim borderTypes As WdBorderType() = {WdBorderType.wdBorderLeft, WdBorderType.wdBorderRight, WdBorderType.wdBorderTop, WdBorderType.wdBorderBottom}
        For Each borderType As WdBorderType In borderTypes
            If table.Borders(borderType).LineStyle = WdLineStyle.wdLineStyleNone Then
                Return False
            End If
        Next
        Return True
    End Function
    Private Sub ResetTable(ByVal table As Object)
        Dim borderTypes As WdBorderType() = {WdBorderType.wdBorderLeft, WdBorderType.wdBorderRight, WdBorderType.wdBorderTop, WdBorderType.wdBorderBottom, WdBorderType.wdBorderHorizontal, WdBorderType.wdBorderVertical}
        For Each borderType As WdBorderType In borderTypes
            With table.Borders(borderType)
                .LineStyle = WdLineStyle.wdLineStyleSingle
                .LineWidth = WdLineWidth.wdLineWidth050pt
            End With
        Next
    End Sub
    Private Sub SetTableFormat(ByVal table As Object)
        Dim borderTypes As WdBorderType() = {WdBorderType.wdBorderLeft, WdBorderType.wdBorderRight, WdBorderType.wdBorderTop, WdBorderType.wdBorderBottom}

        For Each borderType As WdBorderType In borderTypes
            With table.Borders(borderType)
                .LineStyle = WdLineStyle.wdLineStyleSingle
                .LineWidth = WdLineWidth.wdLineWidth100pt
            End With
        Next

        ' 遍历表格的每个单元格
        For Each row As Row In table.Rows
            For Each cell As Cell In row.Cells
                Dim cellText As String = cell.Range.Text.Trim()
                If cellText.Equals(String.Empty) OrElse cellText = Chr(13) OrElse cellText = ChrW(7) Then
                    ' 如果单元格为空，则插入"—"
                    cell.Range.Text = "—"
                ElseIf cellText.Equals("-" & vbCr & ChrW(7)) Then
                    ' 如果单元格内容是"-"，则替换为"—"
                    cell.Range.Text = "—"
                ElseIf cellText.Contains("。") Then
                    ' 找到最后一个句号“。”的位置
                    If cellText.LastIndexOf("。") <> -1 Then
                        ' 删除找到的最后一个句号
                        cell.Range.Text = cellText.Replace(vbCr & ChrW(7), "").Remove(cellText.LastIndexOf("。"), 1)
                    End If
                End If
            Next
        Next

        table.Rows(1).Select()
        Globals.ThisAddIn.Application.Selection.Cells.Borders(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth100pt

        For Each paragraph As Paragraph In table.Range.Paragraphs
            Dim text As String = CType(paragraph.Style, Object).NameLocal
            If Not text.Contains("标准文件_注") AndAlso text <> "标准文件_图表脚注" Then
                paragraph.Range.Font.Size = 9
            End If
        Next

        With table
            .Rows.WrapAroundText = 0
            .RightPadding = 4
            .LeftPadding = 4
        End With

        table.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
        table.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter
    End Sub

    Private Sub Button10_Click(sender As Object, e As RibbonControlEventArgs) Handles Button10.Click
        If Not activedoc() Then Exit Sub
        ' 开启当前文档的修订模式
        currentDoc.TrackRevisions = False
        SearchAndExecuteUnsplit()
    End Sub
    Private Sub SearchAndExecuteUnsplit()
        Dim searchText As String = "（续）" ' 要搜索的特定文字
        Dim found As Boolean = True
        Dim startFrom As Integer = 0 ' 新增变量，表示从文档的哪个位置开始搜索

        While found
            found = False ' 重置 found 标志

            ' 在整个文档中搜索特定文字
            Dim paragraph As Word.Paragraph = Nothing
            For Each para As Word.Paragraph In currentDoc.Paragraphs
                ' 只搜索从 startFrom 位置之后的段落
                If para.Range.Start >= startFrom AndAlso para.Range.Text.Contains(searchText) Then
                    ' 执行一系列操作
                    unsplittab(para)
                    found = True ' 找到并执行操作后设置 found 为 True
                    Exit For ' 退出搜索循环，以便进行下一次搜索
                End If
            Next

            ' 如果没有找到特定文字 "（续）"，则退出循环
            If Not found Then
                Exit While
            End If

            ' 将光标移到文档开头
            Globals.ThisAddIn.Application.Selection.HomeKey(Word.WdUnits.wdStory)
        End While
    End Sub

    Private Sub unsplittab(para As Word.Paragraph)
        ' 将光标移到段落的第一行
        para.Range.Select()
        Globals.ThisAddIn.Application.Selection.HomeKey(Word.WdUnits.wdLine)

        ' 扩展选定区域到段落的最后一行
        Globals.ThisAddIn.Application.Selection.MoveEnd(Word.WdUnits.wdLine, Word.WdMovementType.wdExtend + 1)

        ' 删除选定区域
        Globals.ThisAddIn.Application.Selection.Delete()

        ' 检查段落是否包含表格
        If para.Range.Tables.Count > 0 Then
            ' 示例操作：选择整行
            para.Range.Select()
            Globals.ThisAddIn.Application.Selection.Cells.Borders(Word.WdBorderType.wdBorderTop).LineWidth = Word.WdLineWidth.wdLineWidth025pt
            ResetTable(para.Range.Tables(1))
            SetTableFormat(para.Range.Tables(1))
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As RibbonControlEventArgs) Handles Button9.Click
        If Not activedoc() Then Exit Sub
        ' 切换到单页视图
        Globals.ThisAddIn.Application.ActiveWindow.View.Type = WdViewType.wdPrintView
        SplitTable()
    End Sub
    ' 这个子程序用于处理Microsoft Word中跨页的表格拆分
    Private Sub SplitTable()
        Try
            ' 循环，直到没有跨页的表格
            Do
                ' 检查当前选择是否在表格中
                If Not CType(Globals.ThisAddIn.Application.Selection.Information(Word.WdInformation.wdWithInTable), Boolean) Then
                    MessageBox.Show("请将光标移到待拆分的表格内！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    Exit Do
                End If

                Dim selectedTable As Word.Table = Globals.ThisAddIn.Application.Selection.Tables(1)
                Dim startPageNumber As Integer = CType(Globals.ThisAddIn.Application.Selection.Information(Word.WdInformation.wdActiveEndPageNumber), Integer)
                Dim endPageNumber As Integer = CType(selectedTable.Rows(selectedTable.Rows.Count).Range.Information(Word.WdInformation.wdActiveEndPageNumber), Integer)

                ' 如果选中的表格没有跨页，则退出循环
                If startPageNumber = endPageNumber Then
                    MessageBox.Show("所选表格不再跨页！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    Exit Do
                End If

                ' 复制选中的表格
                selectedTable.Select()

                Dim selection As Word.Selection = Globals.ThisAddIn.Application.Selection
                selection.MoveUp(Missing.Value, 1, Type.Missing)

                ' 得到表格标题文本
                Dim tableTitle As String = GetTableTitle(selection)

                ' 查找跨页的行号
                'Dim splitRowIndex As Integer = FindSplitRowIndex(selectedTable, startPageNumber, endPageNumber)
                Dim splitRowIndex = FindPageBreakRow(selectedTable, startPageNumber)

                ' 拆分表格
                If splitRowIndex <> -1 Then
                    SplitTableAtRow(selectedTable, splitRowIndex)
                    Console.WriteLine("表格从行 " & splitRowIndex.ToString() & " 开始跨页。")
                Else
                    Exit Sub
                    Console.WriteLine("表格没有跨页。")
                End If
                selection.MoveUp(Missing.Value, 1, Type.Missing)
                ' 调整拆分后的表标题格式
                FormatTableTitle(selection, tableTitle)

                ' 向下移动光标到新表格的开始位置
                ' 这里可能需要根据实际情况调整移动的具体方式和距离
                selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=1)
                ' 循环继续，检查下一个表格
            Loop
        Catch ex As Exception
            ' 异常处理
            MessageBox.Show("出现异常：" & ex.Message, "提示")
        End Try
    End Sub

    ' 以下是需要实现的辅助函数
    Private Function GetTableTitle(selection As Word.Selection) As String
        Dim text As String = selection.Range.ListFormat.ListString.Trim() & "　" & selection.Paragraphs(1).Range.Text.Trim()
        text = text.Trim()

        If String.IsNullOrEmpty(text) Then
            text = "上表（续）"
        Else
            text = If(Not text.StartsWith("表"), "上表（续）", text.Replace("（续）", "") & "（续）")
        End If

        Return text
    End Function
    Function FindPageBreakRow(ByVal table As Word.Table, ByVal startPage As Integer) As Integer
        Dim low As Integer = 1
        Dim high As Integer = table.Rows.Count
        Dim currentPage As Integer = startPage

        While low <= high
            Dim mid As Integer = low + (high - low) \ 2
            Dim midPageNumber As Integer = CType(table.Rows(mid).Range.Information(Word.WdInformation.wdActiveEndPageNumber), Integer)

            If midPageNumber = currentPage Then
                ' 继续向下搜索
                low = mid + 1
            Else
                ' 检查是否是跨页的第一行
                Dim previousPageNumber As Integer = CType(table.Rows(mid - 1).Range.Information(Word.WdInformation.wdActiveEndPageNumber), Integer)
                If previousPageNumber = currentPage Then
                    Return mid ' 找到跨页的第一行
                Else
                    high = mid - 1
                End If
            End If
        End While

        Return -1 ' 如果表格没有跨页，返回-1
    End Function

    Private Sub FormatTableTitle(selection As Word.Selection, tableTitle As String)
        selection.TypeText(tableTitle)

        'selection.MoveLeft(Type.Missing, 1, Type.Missing)
        selection.Style = currentDoc.Styles("标准文件_段")

        selection.ParagraphFormat.LineUnitBefore = 0.5F
        selection.ParagraphFormat.LineUnitAfter = 0.5F
        selection.Paragraphs(1).Range.Font.Name = "黑体"

        selection.MoveLeft(Type.Missing, 3, Word.WdMovementType.wdExtend)
        selection.Font.Name = "宋体"

        selection.MoveRight(Type.Missing, 1, Type.Missing)
        selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0.0F
        selection.ParagraphFormat.FirstLineIndent = 0.0F
    End Sub
    Private Sub CopyHeader(ByRef originalTable As Word.Table, ByRef newTable As Word.Table)
        ' 在新表格的顶部插入一行
        newTable.Rows.Add(BeforeRow:=newTable.Rows(1))

        ' 复制原始表格的表头到新表格的第一行
        Dim iCol As Integer
        For iCol = 1 To originalTable.Columns.Count
            newTable.Cell(1, iCol).Range.Text = originalTable.Cell(1, iCol).Range.Text.Replace(vbCr, String.Empty)
        Next
        newTable.Rows(1).Select()
        Globals.ThisAddIn.Application.Selection.Cells.Borders(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth100pt
    End Sub
    Private Function GetTableIndex(ByVal tbl As Word.Table) As Integer
        Dim i As Integer
        For i = 1 To currentDoc.Tables.Count
            If currentDoc.Tables(i).Range.Start = tbl.Range.Start Then
                Return i
            End If
        Next i
        Return 0
    End Function
    Private Sub SplitTableAtRow(selectedTable As Word.Table, rowIndex As Integer)
        Dim originalTableCount As Integer = currentDoc.Tables.Count
        ' 在拆分之前获取当前表格的索引
        Dim originalTableIndex As Integer = GetTableIndex(selectedTable)
        selectedTable.Rows(rowIndex).Select()
        selectedTable.Application.Selection.SplitTable()

        ' 检查是否有新表格被创建
        If currentDoc.Tables.Count > originalTableCount Then
            Dim newTable As Word.Table = currentDoc.Tables(originalTableIndex + 1)
            CopyHeader(selectedTable, newTable)
        End If

    End Sub

    Private Sub Button11_Click(sender As Object, e As RibbonControlEventArgs) Handles Button11.Click
        If Not activedoc() Then Exit Sub
        ' 查找具有特定样式的段落并替换空格，同时删除末尾多余空格
        ProcessFirstPageParagraphs(currentDoc, "标准文件_文件名称", True)

        ' 查找具有特定样式的段落并格式化英文名称，同时删除末尾多余空格
        ProcessFirstPageParagraphs(currentDoc, "封面标准英文名称", False)

        ' 查找具有特定样式的段落并格式化英文名称，同时删除末尾多余空格
        'Com不允许编辑：ProcessFirstPageParagraphs(currentDoc, "标准文件_正文标准名称", False)
    End Sub
    ' 处理指定样式的段落
    Private Sub ProcessFirstPageParagraphs(doc As Word.Document, styleName As String, replaceSpaces As Boolean)
        Dim wordApp As Word.Application = doc.Application
        For Each para As Word.Paragraph In currentDoc.Paragraphs
            If para.Style.NameLocal = "标准文件_章标题" Then Exit Sub
            If para.Style.NameLocal = styleName Then
                wordApp.Selection.Start = para.Range.Start
                wordApp.Selection.End = para.Range.End - 1

                ' 删除末尾多余空格
                wordApp.Selection.Text = Trim(wordApp.Selection.Text)

                If replaceSpaces Then
                    ' 替换一个或多个半角空格为一个全角空格
                    With wordApp.Selection.Find
                        .ClearFormatting()
                        .Text = "[ ]{1,}" ' 使用通配符匹配一个或多个连续的半角空格
                        .Replacement.ClearFormatting()
                        .Replacement.Text = ChrW(12288) ' 全角空格的 Unicode 编码
                        .MatchWildcards = True ' 启用通配符匹配
                        .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                    End With
                    Exit Sub
                Else
                    ' 格式化英文名称
                    wordApp.Selection.Text = LCase(wordApp.Selection.Text)
                    ' 将首个字母转为大写
                    If wordApp.Selection.Text.Length > 0 Then
                        wordApp.Selection.Text = UCase(Mid(wordApp.Selection.Text, 1, 1)) & Mid(wordApp.Selection.Text, 2)
                    End If

                    '' 第一步：替换所有的 "-" 为 "—"
                    'With wordApp.Selection.Find
                    '    .ClearFormatting()
                    '    .Text = "-"
                    '    .Replacement.ClearFormatting()
                    '    .Replacement.Text = ""
                    '    .Execute(Replace:=Word.WdReplace.wdReplaceNone)
                    '    While .Found
                    '        ' 检查前后是否有空格
                    '        Dim beforeChar As String = ""
                    '        If wordApp.Selection.Start > 1 Then
                    '            beforeChar = doc.Range(wordApp.Selection.Start - 1, wordApp.Selection.Start).Text
                    '        End If
                    '        Dim afterChar As String = ""
                    '        If wordApp.Selection.End <doc.Content.End Then
                    '            afterChar = doc.Range(wordApp.Selection.End, wordApp.Selection.End + 1).Text
                    '        End If

                    '        ' 根据需要添加空格
                    '        Dim replacementText As String = "—"
                    '        If beforeChar <> " " Then replacementText = " " & replacementText
                    '        If afterChar <> " " Then replacementText = replacementText & " "

                    '        wordApp.Selection.Text = replacementText

                    '        ' 移动到下一个 "-" 以避免重复替换
                    '        wordApp.Selection.Start = wordApp.Selection.Start + replacementText.Length
                    '        wordApp.Selection.End = wordApp.Selection.Start
                    '        .Execute(Replace:=Word.WdReplace.wdReplaceNone)
                    '    End While
                    'End With

                    ' 第二步：将 "—","-"或":" 前后空格去掉,之后的英文字母转为大写
                    ' 获取当前段落的范围
                    Dim currentParagraph As Word.Range = wordApp.Selection.Paragraphs(1).Range
                    Dim paragraphStart As Integer = currentParagraph.Start
                    Dim paragraphEnd As Integer = currentParagraph.End

                    With wordApp.Selection.Find
                        .ClearFormatting()
                        .Text = "[-—:]" ' 将 "-" 放在开头，匹配 "—"、"-" 或 ":"
                        .MatchWildcards = True

                        ' 设置查找范围为当前段落
                        wordApp.Selection.SetRange(paragraphStart, paragraphEnd)

                        While .Execute(Replace:=Word.WdReplace.wdReplaceNone) And wordApp.Selection.Start <= paragraphEnd
                            ' 检查并删除前面的空格
                            If wordApp.Selection.Start > paragraphStart Then
                                Dim beforeCharRange As Word.Range = doc.Range(wordApp.Selection.Start - 1, wordApp.Selection.Start)
                                If beforeCharRange.Text = " " Then
                                    beforeCharRange.Delete()
                                    ' 更新段落结束位置
                                    paragraphEnd = currentParagraph.End
                                End If
                            End If

                            ' 检查并删除后面的空格
                            If wordApp.Selection.End < paragraphEnd Then
                                Dim afterCharRange As Word.Range = doc.Range(wordApp.Selection.End, wordApp.Selection.End + 1)
                                If afterCharRange.Text = " " Then
                                    afterCharRange.Delete()
                                    ' 更新段落结束位置
                                    paragraphEnd = currentParagraph.End
                                Else
                                    ' 如果后面不是空格，移动选择范围以选中并转换为大写
                                    wordApp.Selection.SetRange(wordApp.Selection.End, wordApp.Selection.End + 1)
                                    wordApp.Selection.Text = UCase(wordApp.Selection.Text)
                                End If
                            End If

                            ' 移动到下一个匹配的字符，但不超过段落范围
                            wordApp.Selection.SetRange(wordApp.Selection.End, paragraphEnd)
                        End While
                    End With
                    Exit Sub
                End If
            End If
        Next
    End Sub


    Private Sub ProcessParagraphWithStyle2(doc As Word.Document, styleName As String, replaceSpaces As Boolean)
        For Each para As Word.Paragraph In doc.Paragraphs
            If para.Style.NameLocal = styleName Then
                Dim range As Word.Range = para.Range

                ' 删除末尾多余空格
                range.Text = Trim(range.Text)

                If replaceSpaces Then
                    ' 替换半角空格为全角空格
                    range.Find.Text = " "
                    range.Find.Replacement.Text = ChrW(12288) ' 全角空格的 Unicode 编码
                    range.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
                Else
                    ' 格式化英文名称
                    ' 全部转为小写
                    range.Text = LCase(range.Text)

                    ' 将首个字母转为大写
                    If range.Text.Length > 0 Then
                        range.Text = UCase(Mid(range.Text, 1, 1)) & Mid(range.Text, 2)
                    End If

                    ' 替换 "-" 为 "—" 并将其后的首个字母转为大写
                    While range.Find.Execute(FindText:="-", ReplaceWith:="—", Replace:=Word.WdReplace.wdReplaceOne)
                        If range.Start < range.End Then
                            Dim nextCharRange As Word.Range = doc.Range(range.End, range.End + 1)
                            nextCharRange.Text = UCase(nextCharRange.Text)
                        End If
                    End While
                End If
            End If
        Next
    End Sub

    Private Sub Button12_Click(sender As Object, e As RibbonControlEventArgs) Handles Button12.Click
        If Not activedoc() Then Exit Sub
        Dim introStarted As Boolean = False
        Dim introEnded As Boolean = False

        For Each para In currentDoc.Paragraphs
            Dim paraStyle As String = para.Style.NameLocal
            Dim paraText As String = para.Range.Text

            ' 检查引言部分的开始
            If paraStyle = "标准文件_前言、引言标题" And paraText.Contains("引言") Then
                introStarted = True
                introEnded = False
            End If

            ' 检查引言部分的结束
            If paraStyle = "标准文件_正文标准名称" Then
                introEnded = True
            End If

            ' 在引言部分中检查是否包含不允许的词汇
            If introStarted And Not introEnded Then
                If paraText.Contains("应") Or paraText.Contains("不应") Then
                    ' 在有问题的引言标题上添加批注
                    currentDoc.Comments.Add(para.Range, "可能存在要求性的表述。")
                    Exit For ' 找到第一个实例后停止检查
                End If
            End If
        Next
    End Sub

    Private Sub Button13_Click(sender As Object, e As RibbonControlEventArgs) Handles Button13.Click
        If Not activedoc() Then Exit Sub
        Dim chapterStarted As Boolean = False
        Dim chapterStartIndex As Integer = 0
        Dim paragraphsDetails As New List(Of Tuple(Of String, String))() ' 存储段落首字符和完整文本

        Dim paraCount As Integer = currentDoc.Paragraphs.Count
        Dim i As Integer = 1

        While i <= paraCount
            Dim para As Word.Paragraph = currentDoc.Paragraphs(i)
            Dim paraStyle As String = para.Style.NameLocal
            Dim paraText As String = para.Range.Text.Trim()

            ' 检查章节的开始
            If Not chapterStarted AndAlso paraStyle = "标准文件_章标题" AndAlso paraText.Contains("缩略语") Then
                chapterStarted = True
                chapterStartIndex = i
                i += 2 ' 跳过章节标题下的第一段
                Continue While
            End If

            ' 如果在“缩略语”章节内
            If chapterStarted Then
                ' 收集段落首字符和文本，准备排序
                If Not String.IsNullOrEmpty(paraText) Then
                    Dim match As Match = Regex.Match(paraText, "^[1-9A-Za-z]+")
                    If match.Success AndAlso Not IsStringPresentTimes(match.Value, 1) Then
                        currentDoc.Comments.Add(currentDoc.Paragraphs(chapterStartIndex).Range, "术语" & match.Value & "在文中没有出现。")
                    End If
                    paragraphsDetails.Add(Tuple.Create(paraText, paraText))
                End If

                ' 删除当前段落
                para.Range.Delete()
                paraCount -= 1

                ' 检查章节的结束
                If i <= paraCount Then
                    Dim nextParaStyle As String = currentDoc.Paragraphs(i).Style.NameLocal
                    If nextParaStyle = "标准文件_章标题" Then
                        ' 按段落首字符排序
                        paragraphsDetails.Sort(Function(x, y) x.Item2.CompareTo(y.Item2))
                        ' 使用LINQ将所有段落文本合并成一个字符串，段落之间以换行符分隔
                        Dim combinedText As String = String.Join(vbCrLf, paragraphsDetails.Select(Function(detail) detail.Item2))

                        ' 在章节开头下面插入排好序的内容
                        Dim insertRange As Word.Range = currentDoc.Paragraphs(chapterStartIndex + 1).Range

                        insertRange.InsertAfter(combinedText & vbCr)
                        insertRange.Style = currentDoc.Styles("标准文件_段")

                        ' 退出函数
                        Exit Sub
                    End If
                End If
            Else
                i += 1
            End If
        End While
    End Sub

    'Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
    '    If Not activedoc() Then Exit Sub

    '    ' 从用户获取需要查找的字符或字符串
    '    Dim userInput As String = InputBox("请输入需要查找的字符或字符串:", "查找字符", "x")
    '    If String.IsNullOrEmpty(userInput) Then Exit Sub

    '    ' 为用户输入构建正则表达式
    '    Dim pattern As String = $"(?<=[\u4e00-\u9fff]|[，。！？,.\s])({System.Text.RegularExpressions.Regex.Escape(userInput)})(?=[\u4e00-\u9fff]|[，。！？,.\s])"
    '    Dim regex As New System.Text.RegularExpressions.Regex(pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

    '    ' 遍历文档中的所有内容，查找匹配项
    '    Dim matches As System.Text.RegularExpressions.MatchCollection = regex.Matches(currentDoc.Content.Text)
    '    Console.WriteLine($"找到匹配项数量: {matches.Count}") ' 输出匹配项的数量，便于调试

    '    For Each match As System.Text.RegularExpressions.Match In matches
    '        Dim matchRange As Range = currentDoc.Range(Start:=match.Index, End:=match.Index + match.Length)
    '        ' 应用斜体和Cambria字体
    '        With matchRange.Font
    '            .Italic = True
    '            .Name = "Cambria"
    '        End With
    '    Next
    'End Sub

    'Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
    '    If Not activedoc() Then Exit Sub

    '    ' 从用户获取需要查找的字符或字符串
    '    Dim userInput As String = InputBox("请输入需要查找的字符或字符串:", "查找字符", "x")
    '    If String.IsNullOrEmpty(userInput) Then Exit Sub

    '    ' 为用户输入构建正则表达式
    '    Dim pattern As String = $"(?<=[\u4e00-\u9fff]|[，。！？,.\s])({System.Text.RegularExpressions.Regex.Escape(userInput)})(?=[\u4e00-\u9fff]|[，。！？,.\s])"
    '    Dim regex As New System.Text.RegularExpressions.Regex(pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

    '    ' 遍历文档中的所有内容，查找匹配项
    '    Dim currentRange As Range = currentDoc.Content
    '    Dim matches As System.Text.RegularExpressions.MatchCollection = regex.Matches(currentRange.Text)

    '    For Each match As System.Text.RegularExpressions.Match In matches
    '        MessageBox.Show($"找到匹配项: {match.Value}") ' 显示调试信息
    '        ' 在这里，我们需要一个方法来定位Word中的匹配项并应用格式，但由于直接操作Word文档进行格式化比较复杂，以下代码是概念性的
    '        ' 可以使用currentRange.SetRange等方法来定位和格式化匹配的文本，但需要精确计算位置
    '    Next

    '    ' 注意：此处的代码并未直接应用格式，因为需要更复杂的逻辑来定位文本并应用格式
    'End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If Not activedoc() Then Exit Sub
        ' 开启当前文档的修订模式
        currentDoc.TrackRevisions = True

        ' 从用户获取需要查找的字符或字符串
        Dim userInput As String = InputBox("请输入需要查找的字符或字符串:", "查找字符", "x")
        If String.IsNullOrEmpty(userInput) Then Exit Sub
        Dim modifiedCount = 0

        ' 正则表达式用于匹配非英文字符的边界
        Dim pattern As String = $"(?<=[\u4e00-\u9fff]|[，。！？,.\s])(?<![a-zA-Z])({System.Text.RegularExpressions.Regex.Escape(userInput)})(?![a-zA-Z])(?=[\u4e00-\u9fff]|[，。！？,.\s])"
        Dim regex As New System.Text.RegularExpressions.Regex(pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        ' 遍历文档中的所有内容，查找匹配项
        Dim currentRange As Word.Range = currentDoc.Content
        currentRange.Find.ClearFormatting()
        With currentRange.Find
            .Text = userInput
            .MatchCase = True
            .MatchWholeWord = False  ' 修改这里，以匹配单个字符
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With

        Do While currentRange.Find.Execute
            Dim isSuperscriptOrSubscript As Boolean = currentRange.Font.Superscript <> 0 Or currentRange.Font.Subscript <> 0
            Dim isValidContext As Boolean = True

            ' 检查字符前后是否有英文字母，除非是上标或下标
            If Not isSuperscriptOrSubscript Then
                Dim currentStart As Integer = currentRange.Start
                Dim currentEnd As Integer = currentRange.End
                Dim isAlphabeticBefore As Boolean = currentStart > 1 AndAlso System.Text.RegularExpressions.Regex.IsMatch(currentDoc.Range(currentStart - 1, currentStart).Text, "[a-zA-Z]")
                Dim isAlphabeticAfter As Boolean = currentEnd < currentDoc.Content.End AndAlso System.Text.RegularExpressions.Regex.IsMatch(currentDoc.Range(currentEnd, currentEnd + 1).Text, "[a-zA-Z]")
                isValidContext = Not isAlphabeticAfter And Not isAlphabeticBefore
            End If

            If isValidContext Then
                With currentRange.Font
                    .Italic = True
                    .Name = "Times New Roman"
                End With
                modifiedCount += 1
            End If

            ' 确保不会超出文档范围
            If currentRange.End + 1 <= currentDoc.Content.End Then
                currentRange.Start = currentRange.End + 1
            Else
                Exit Do
            End If
        Loop
        ' 显示消息框
        If modifiedCount > 0 Then
            MessageBox.Show(modifiedCount & "个变量 " & userInput & " 已被改为斜体，请检查确认。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("没有发现满足条件的变量 " & userInput & "。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As RibbonControlEventArgs) Handles Button15.Click
        If Not activedoc() Then Exit Sub
        Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim progressForm As New ProgressForm()
        Dim wordAppHwnd As IntPtr = New IntPtr(CType(wordApp.ActiveWindow.Hwnd, Long))
        Dim dpi As Integer = NativeMethods.GetDpiForWindow(wordAppHwnd)
        'Dim graphics As Graphics = Graphics.FromHwnd(IntPtr.Zero)
        Dim scaleX As Single = dpi / 96.0 'graphics.DpiX / 96 ' 默认DPI为96
        Dim scaleY As Single = dpi / 96.0 'graphics.DpiY / 96
        'Graphics.Dispose()

        Dim wordWindowRectangle As Drawing.Rectangle = GetWordWindowRectangle(wordApp)

        ' 计算窗体的新位置，使之居中于Word窗口
        Dim formLeft As Integer = wordWindowRectangle.Left + (wordWindowRectangle.Width - progressForm.Width * scaleX) / 2
        Dim formTop As Integer = wordWindowRectangle.Top + (wordWindowRectangle.Height - progressForm.Height * scaleY) / 2

        progressForm.StartPosition = FormStartPosition.Manual
        progressForm.Location = New Drawing.Point(formLeft, formTop)

        ' 显示窗体
        progressForm.Show()
        wordApp.ScreenUpdating = False

        ' 不需要更改背景颜色的样式名称集合
        Dim excludedStyles As New List(Of String) From {"章", "条", "标题", "附录", "图", "表", "注"}
        Dim isInList As Boolean = False
        Dim lastParagraphStyle As String = String.Empty
        Dim lastParagraphEndsWith As Char = Char.MinValue
        Dim listEndsWith As Char = Char.MinValue ' 记录列项应该以什么符号结尾
        Dim firstListItem As Boolean = True ' 标记是否为列项的第一个段落
        Dim listItemEndCharacter As Char = Char.MinValue ' 记录除最后一个列项外，其他列项应使用的标点符号
        Dim missLeading As Boolean = False

        Dim totalParagraphs As Integer = currentDoc.Paragraphs.Count
        Dim currentParagraph As Integer = 0
        Dim lastProgress As Integer = 0
        Dim currentProgress As Integer = 0
        Try
            For Each para As Word.Paragraph In currentDoc.Paragraphs
                currentParagraph += 1
                currentProgress = CInt((currentParagraph / totalParagraphs) * 100)

                ' 仅当进度实际改变时更新UI
                If currentProgress <> lastProgress Then
                    ' 更新进度条
                    progressForm.UpdateProgress(currentProgress)
                    progressForm.UpdateMessage("进度：" & currentProgress & "%")
                    lastProgress = currentProgress
                End If
                ' 减少DoEvents调用，只在关键进度更新时调用
                If currentParagraph Mod 10 = 0 Then
                    System.Windows.Forms.Application.DoEvents()
                End If

                '' 判断该段落前面是否有编号且样式名称不包含在排除列表中
                'If para.Range.ListFormat.ListType <> Word.WdListType.wdListNoNumbering AndAlso Not excludedStyles.Any(Function(style) para.Range.Style.NameLocal.Contains(style)) Then
                '    ' 将有编号的段落的背景颜色设置为黄色
                '    para.Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorYellow
                'End If

                Dim currentStyle As String = para.Style.NameLocal
                Dim currentText As String = para.Range.Text.Trim()
                If currentText.Length > 0 Then
                    If para.Range.ListFormat.ListType <> Word.WdListType.wdListNoNumbering AndAlso Not excludedStyles.Any(Function(style) currentStyle.Contains(style)) Then
                        '替换末尾英文标点符号为中文标点符号
                        ReplaceEnglishPunctuationAtEndOfParagraph(currentDoc, para)
                        lastParagraphEndsWith = ReplaceEnglishPunctuationAtEndOfParagraph(currentDoc, para.Previous)

                        If Not isInList Then
                            ' 检查上一段落是否符合特定样式和结尾符号
                            If lastParagraphStyle.Contains("段") AndAlso (lastParagraphEndsWith = "："c OrElse lastParagraphEndsWith = "。"c) Then
                                isInList = True ' 进入列项
                                listItemEndCharacter = If(lastParagraphEndsWith = "："c, "；"c, "。"c) ' 决定列项结束符
                                firstListItem = True
                            ElseIf Not missLeading Then
                                para.Range.Comments.Add(para.Range, "缺少引导语（GB/T 1.1—2020的7.5）或前段样式不对。")
                                missLeading = True
                            End If
                        End If

                        ' 处理列项段落的标点符号
                        If isInList Then
                            ' 对于列项的第一个段落，检查是否遵循了正确的标点规则
                            If firstListItem Then
                                '对于第一段如果不期望是分号的情况下是逗号，那么纠正
                                If listItemEndCharacter = "；"c And currentText.EndsWith("，"c) Then
                                    listItemEndCharacter = "，"c
                                End If

                                If Not currentText.EndsWith(listItemEndCharacter) Then
                                    'Dim commentText As String = $"列项的开始段落的结尾符号应为'{listItemEndCharacter}'。"
                                    'para.Range.Comments.Add(para.Range, commentText)
                                    ReplaceLastCharacterInParagraphUsingSelection(currentDoc, para, listItemEndCharacter)
                                End If
                                firstListItem = False
                            Else
                                ' 检查下一个段落是否还属于列表
                                Dim isLastListItem As Boolean
                                If para.Next IsNot Nothing Then
                                    ' 检查下一个段落的样式是否在排除列表中，或者下一个段落是否没有编号
                                    isLastListItem = para.Next.Range.ListFormat.ListType = Word.WdListType.wdListNoNumbering OrElse excludedStyles.Any(Function(style) para.Next.Range.Style.NameLocal.Contains(style))
                                Else
                                    ' 如果没有下一个段落，则当前段落必定是列表中的最后一段
                                    isLastListItem = True
                                End If

                                ' 根据是否是最后一个列表项来决定结束字符
                                Dim expectedChar As Char = If(isLastListItem, "。"c, listItemEndCharacter)

                                If Not currentText.EndsWith(expectedChar) And Not isLastListItem Then
                                    'Dim commentText As String = $"列项的结尾符号应为'{expectedChar}'。"
                                    'para.Range.Comments.Add(para.Range, commentText)
                                    ReplaceLastCharacterInParagraphUsingSelection(currentDoc, para, expectedChar)
                                End If
                            End If
                        End If
                    ElseIf isInList Then
                        ' 当前段落不是列项，且之前已经进入列项
                        isInList = False
                        firstListItem = True
                        missLeading = False
                    End If
                End If

                ' 更新上一段落的样式和结尾符号
                lastParagraphStyle = currentStyle
                If currentText.Length > 0 Then lastParagraphEndsWith = currentText.Last()
            Next
            wordApp.ScreenUpdating = True
        Catch ex As Exception
            MessageBox.Show("发生异常：" & ex.Message, "异常通知", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 完成处理后关闭进度窗口
            progressForm.Close()
        End Try
    End Sub
    Public Function ReplaceEnglishPunctuationAtEndOfParagraph(ByVal doc As Document, ByVal paragraph As Paragraph) As String
        ' 确保段落不为空
        If paragraph IsNot Nothing Then
            ' 获取段落的最后一个字符
            Dim lastChar As String = paragraph.Range.Text.Substring(paragraph.Range.Text.Length - 2, 1)

            ' 准备替换规则：英文标点到中文标点的映射
            Dim replaceRules As New Dictionary(Of String, String) From {
            {".", "。"},
            {":", "："},
            {";", "；"},
            {",", "，"}
        }

            ' 检查是否需要替换
            If replaceRules.ContainsKey(lastChar) Then
                ' 执行替换
                doc.Application.Selection.Start = paragraph.Range.End - 2
                doc.Application.Selection.End = paragraph.Range.End - 1
                doc.Application.Selection.Text = replaceRules(lastChar)
                lastChar = doc.Application.Selection.Text
            End If
            Return lastChar
        End If
    End Function

    Public Sub ReplaceLastCharacterInParagraphUsingSelection(ByVal doc As Document, ByVal paragraph As Paragraph, ByVal replacementChar As Char)
        Dim chinesePunctuation As String = "。，、；：？！（）【】《》「」『』"
        ' 确保段落不为空
        If paragraph IsNot Nothing Then
            doc.Application.Selection.Start = paragraph.Range.End - 2
            doc.Application.Selection.End = paragraph.Range.End - 1

            ' 再次确认是否选中了文本
            If Not String.IsNullOrEmpty(doc.Application.Selection.Text) AndAlso chinesePunctuation.Contains(doc.Application.Selection.Text) Then
                ' 替换选中的最后一个字符
                doc.Application.Selection.Text = replacementChar
            End If
        End If
    End Sub


    Private Function GetWordWindowRectangle(wordApp As Application) As Drawing.Rectangle
        ' 获取Word窗口的位置和大小
        With wordApp.ActiveWindow
            Return New Drawing.Rectangle(.Left, .Top, .Width, .Height)
        End With
    End Function

    Private Sub searchStd_Click(sender As Object, e As RibbonControlEventArgs) Handles searchStd.Click
        Dim dialog As New BibsearchDialog()
        dialog.ShowDialog()
    End Sub
End Class
