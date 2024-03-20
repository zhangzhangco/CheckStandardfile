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
Imports Microsoft.Office
Imports System.Diagnostics.Eventing.Reader
Imports System.IO
Imports System.Security.Cryptography
Imports Microsoft.Office.Core
Imports Newtonsoft.Json
Imports System.Windows.Input

Public Class Ribbon1
    Private currentDoc As Word.Document
    ' 使用静态HttpClient实例以提高效率和资源复用
    Public Shared ReadOnly HttpClientInstance As New HttpClient()
    ' 获取当前Word应用程序和活动文档的引用
    Dim wordApp As Microsoft.Office.Interop.Word.Application ' = Globals.ThisAddIn.Application
    ' 创建ProgressHandler实例，用于管理进度条
    Dim progressHandler As New ProgressHandler()
    Public Property originalDocPath As String
    Public Property tempPathOriginal As String
    Public Property tempPathModified As String
    Public Property LicenseKey As String
    Public Property Llm As String
    Public Property LlmKey As String
    Public Property IsVip As Boolean = False
    Public Shared ReadOnly InstalledPath = Environment.GetEnvironmentVariable("APPDATA") & "\RelatonChina\标准形式检查助手\"
    Public Shared ReadOnly IniPath = InstalledPath & "setting.ini"
    Public Shared ReadOnly StylePath = InstalledPath & "行业标准.dotx"
    Public Shared ReadOnly UpdaterPath = InstalledPath & "更新.exe"

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        wordApp = Globals.ThisAddIn.Application
        LoadSettings()
    End Sub

    Private Sub About_Click(sender As Object, e As RibbonControlEventArgs) Handles AboutBtn.Click
        Dim aboutMessage As String = "形式检查助手" & Environment.NewLine
        aboutMessage &= "版本: 0.2.2" & Environment.NewLine
        aboutMessage &= "WeChat：HelloLLM2035" & Environment.NewLine
        aboutMessage &= "用于辅助进行标准形式检查和编制的小工具。"

        MessageBox.Show(aboutMessage, "关于", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub StructureChk_Click(sender As Object, e As RibbonControlEventArgs) Handles StructureChkBtn.Click
        Dim para As Paragraph
        Dim nextPara As Word.Paragraph
        Dim currentLevel As Integer
        Dim nextLevel As Integer
        Dim subLevelPara As Word.Paragraph

        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If
        ' 开启当前文档的修订模式
        'DecryptDoc(currentDoc)
        'currentDoc.TrackRevisions = True

        progressHandler.ProgressStartWaiting()
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
        progressHandler.ProgressEnd()
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

    Private Sub BibValid_Click(sender As Object, e As RibbonControlEventArgs) Handles BibValidBtn.Click
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If
        ' 开启当前文档的修订模式
        'If (sender Is Nothing) Then
        '    currentDoc.TrackRevisions = False
        'Else
        '    DecryptDoc(currentDoc)
        '    currentDoc.TrackRevisions = True
        'End If

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

        progressHandler.ProgressStartWaiting()
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
                        fileNames.Add(StandardDocument.FormatedFileName(para.Range.Text))
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
                    Dim stdfilename = arrFileNames(j)
                    '尊享模式才有
                    If (sender Is Nothing) Then
                        '构建类，拆解文件名，得到标准编号和标准名称
                        Dim stddoc = New StandardDocument(stdfilename)
                        Dim output As String
                        '验证标准引用的有效性，已废止做批注，已更新替换标准名
                        If stddoc.isDomestic Then
                            output = SearchCnGovStd(stddoc.Code, HttpClientInstance)
                        Else
                            '仅对支持的三大国际标准进行查询
                            If String.IsNullOrEmpty(stddoc.Code) Then
                                output = String.Empty
                            Else
                                output = SearchInterStd(stddoc.Code, HttpClientInstance, LicenseKey)
                            End If
                        End If
                        '查到了就添加回车
                        If Not String.IsNullOrEmpty(output) Then
                            stdfilename = output & vbCrLf
                        End If

                        insertPoint.Text = stdfilename
                        insertPoint.Style = "标准文件_段"
                        '属于三大标准，但没有查到
                        If Not String.IsNullOrEmpty(stddoc.Code) AndAlso String.IsNullOrEmpty(output) Then
                            currentDoc.Comments().Add(insertPoint, $"{stddoc.Code} 不存在或者已废止。")
                        End If

                        insertPoint = currentDoc.Range(insertPoint.End, insertPoint.End)
                    Else
                        insertPoint.Text = stdfilename
                        insertPoint.Style = "标准文件_段"
                        insertPoint = currentDoc.Range(insertPoint.End, insertPoint.End)
                    End If
                End If
            Next j
        Else
            If Not sender Is Nothing Then
                MsgBox("未找到'规范性引用文件'章节，您可能未使用SET 2020编写此文件。")
            End If
            '批注缺少章节
            missingBib(currentDoc)
        End If
        progressHandler.ProgressEnd()
    End Sub
    Private Sub missingBib(doc As Document)
        ' 遍历文档中的每个段落
        For Each para As Word.Paragraph In doc.Paragraphs
            ' 检查段落的样式名称和内容
            If para.Style.NameLocal = "标准文件_章标题" AndAlso para.Range.Text.Contains("范围") Then
                ' 在满足条件的段落中添加批注
                para.Range.Comments.Add(para.Range, "此章后必须有‘规范性引用文件’一章")
            End If
        Next
    End Sub
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

    Private Sub TermsChk_Click(sender As Object, e As RibbonControlEventArgs) Handles TermsChkBtn.Click
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If
        Dim para As Word.Paragraph

        ' 开启当前文档的修订模式
        currentDoc.TrackRevisions = False

        progressHandler.ProgressStartWaiting()
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
                        '第一个空格全角，后面的英文小写
                        ReplaceWithQuanjiaoAndConditionallyLowercase(para)
                        Dim tt = GetLeadingText(para)
                        If Not IsStringPresentTimes(tt, 2) Then
                            currentDoc.Comments.Add(para.Range, "'" & tt & "'在文中出现少于两次，应从本章中移除。")
                        End If
                    End If
                    ' 移至下一个段落
                    para = nextPara
                End While
            End If
        Next para
        progressHandler.ProgressEnd()
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
    Private Sub ReplaceWithQuanjiaoAndConditionallyLowercase(para As Word.Paragraph)
        If para Is Nothing Then Exit Sub
        If Not String.IsNullOrWhiteSpace(para.Range.Text.Trim) Then
            Dim range As Word.Range = para.Range
            range.SetRange(Start:=para.Range.Start, End:=para.Range.End - 1) ' 避开段落末尾的特殊字符

            ' 统一处理全角空格和半角空格，将全角空格替换为半角空格
            Dim text As String = vbVerticalTab & range.Text.Replace("　", " ").TrimStart(vbVerticalTab, " ") ' 全角空格替换为半角空格

            ' 分割段落文本为两部分：第一个空格前的文本和第一个空格后的文本
            Dim parts() As String = Split(text, " ", 2)
            If parts.Length > 1 Then
                Dim firstPart As String = parts(0) ' 可能包含缩略语的部分
                Dim secondPart As String = parts(1) ' 第一个空格后的部分，可能包含英文单词

                ' 从第一部分中提取所有大写英文字符串
                Dim regex As New Regex("[A-Z]+")
                Dim uppercaseWord As String = ""
                Dim match As Match = regex.Match(firstPart)
                If match.Success Then
                    uppercaseWord = match.Value ' 提取到的大写英文字符串
                End If

                ' 对第二部分的英文单词进行处理，保留特定的大写英文字符串
                Dim secondPartProcessed As String = Regex.Replace(secondPart, "\b[A-Za-z]+\b", Function(m)
                                                                                                   If m.Value.ToUpper() = uppercaseWord Then
                                                                                                       Return uppercaseWord ' 保持特定的大写字符串
                                                                                                   Else
                                                                                                       Return m.Value.ToLower() ' 其他转换为小写
                                                                                                   End If
                                                                                               End Function)

                ' 重组段落文本
                range.Text = firstPart & "　" & secondPartProcessed
            End If

            ' 应用样式和格式
            range.Style = "标准文件_术语条一"
            range.Font.Name = "黑体"
            range.ParagraphFormat.LeftIndent = 24
            range.ParagraphFormat.CharacterUnitFirstLineIndent = -2
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

    Private Sub BibRefChk_Click(sender As Object, e As RibbonControlEventArgs) Handles BibRefChkBtn.Click
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If
        Dim reftext As String

        ' 开启当前文档的修订模式
        'DecryptDoc(currentDoc)
        'currentDoc.TrackRevisions = True

        Dim regEx As Regex
        regEx = New Regex("(([A-Z]{2,})([_/])([A-Z])\s([0-9]{1,5}(?:\.[0-9]{1,3})?)([-—])([0-9]{4}))|(([A-Z]{2.})\s([0-9]+)(?:([-])?([0-9]))(:[0-9]{4})?)")

        reftext = extracteChapterText("规范性引用文件") & extracteChapterText("参考文献")

        progressHandler.ProgressStartWaiting()
        ProcessParagraphs(currentDoc, regEx, reftext)
        ProcessTables(currentDoc, regEx, reftext)
        progressHandler.ProgressEnd()
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
        doc = currentDoc

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

    Private Sub BignumMdf_Click(sender As Object, e As RibbonControlEventArgs) Handles BignumMdfBtn.Click
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If
        DecryptDoc(currentDoc)
        currentDoc.TrackRevisions = True ' 开启修订模式

        Dim regEx As Regex
        regEx = New Regex("\b(?<![a-zA-Z\d .:/\-—""]) \ d{5,}(?:\.\d+)?\b(?![\-/])", RegexOptions.IgnoreCase)
        progressHandler.ProgressStartWaiting()
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
        progressHandler.ProgressEnd()
    End Sub


    Function FormatNumberWithCommas(num As String) As String
        Dim parts As String() = num.Split("."c)
        Dim integerPart As String = parts(0)
        Dim decimalPart As String = If(parts.Length > 1, "." + parts(1), "")

        Dim regex As Regex = New Regex("(\d)(?=(\d{3})+(?!\d))")
        integerPart = regex.Replace(integerPart, "$1,")

        Return integerPart + decimalPart
    End Function

    Private Sub UnitMdf_Click(sender As Object, e As RibbonControlEventArgs) Handles UnitMdfBtn.Click
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If
        ' 开启当前文档的修订模式
        DecryptDoc(currentDoc)
        currentDoc.TrackRevisions = True
        ' 保存当前的修订视图状态
        Dim originalShowRevisions As Boolean = wordApp.ActiveWindow.View.ShowRevisionsAndComments

        ' 设置为最终状态视图，以隐藏修订内容
        wordApp.ActiveWindow.View.ShowRevisionsAndComments = False

        Dim unitPairs As String
        unitPairs = "米|m|千克|kg|秒|s|安培|A|毫安|mA|摩尔|mol|坎德拉|cd|" &
                    "牛顿|N|焦耳|J|瓦特|W|帕斯卡|Pa|伏特|V|欧姆|Ω|库仑|C|" &
                    "法拉|F|特斯拉|T|亨利|H|赫兹|Hz|勒克斯|lx|摄氏度|℃|升|l|" &
                    "克|g|毫米|mm|厘米|cm|千米|km|毫克|mg|微克|μg|吨|t|" &
                    "毫秒|ms|微秒|μs|纳秒|ns|分钟|min|小时|h|" &
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
                    "千伏安|kVA|兆伏安|MVA|吉伏安|GVA|公里每小时|km/h|纳米|nm|微米|μm|度|°|比特|bit"
        Dim unitMap As New Dictionary(Of String, String) From {
                                {"米", "m"},
                                {"千克", "kg"},
                                {"秒", "s"},
                                {"安培", "A"},
                                {"毫安", "mA"},
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
                                {"微米", "μm"},
                                {"度", "°"},
                                {"比特", "bit"}
                            }

        Dim regEx As Regex
        'regEx = New Regex("([-+]?\d*\.?\d+\/?\d*)\s?(" & unitPairs & ")")
        regEx = New Regex("(?<![A-Za-z:_.])([-+]?\d*\.?\d+\/?\d*)\s?(" & unitPairs & ")(?![A-Za-z1-9_.:+-])")
        progressHandler.ProgressStartWaiting()

        ' 替换㎡为m²
        ReplaceTextWithSuperscript(currentDoc, "㎡", "m", "2")

        ' 替换m³为m³
        ReplaceTextWithSuperscript(currentDoc, "m³", "m", "3")

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
                newText = match.Groups(1).Value & ChrW(&H2005) & unit
                If unit = "°" Then 'unit = "℃" Or unit = "℉" Or 
                    newText = match.Groups(1).Value & unit
                End If

                ' 检查原文本与新文本是否不同，若不同，则替换
                If originalText <> newText Then
                    ' 先检查是否以"2"或"3"结尾
                    If newText.EndsWith("2") Or newText.EndsWith("3") Then
                        ' 先设置除最后一个字符外的所有文本
                        Dim textWithoutLastChar As String = newText.Substring(0, newText.Length - 1)
                        matchRange.Text = textWithoutLastChar

                        ' 为最后一个字符创建新的Range
                        Dim lastCharRange As Word.Range = currentDoc.Range(Start:=matchRange.End, End:=matchRange.End)
                        lastCharRange.Text = newText.Substring(newText.Length - 1)

                        ' 应用上标格式
                        lastCharRange.Font.Superscript = True
                    Else
                        ' 如果不以"2"或"3"结尾，正常替换文本
                        matchRange.Text = newText
                    End If
                End If
            Next
        Next
        progressHandler.ProgressEnd()

        ' 恢复原始的修订视图状态
        wordApp.ActiveWindow.View.ShowRevisionsAndComments = originalShowRevisions
    End Sub

    Private Sub ReplaceTextWithSuperscript(doc As Word.Document, searchText As String, baseText As String, superText As String)
        Dim rng As Word.Range = doc.Content

        rng.Find.ClearFormatting()
        rng.Find.Text = searchText
        rng.Find.Replacement.ClearFormatting()

        While rng.Find.Execute(FindText:=searchText, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceNone)
            rng.Text = baseText
            rng.Font.Superscript = 0 ' 先清除可能存在的上标格式
            rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

            ' 插入上标文本
            Dim superRng As Word.Range = doc.Range(rng.Start, rng.Start)
            superRng.Text = superText
            superRng.Font.Superscript = 1

            rng.Start = superRng.End
            rng.End = doc.Content.End
        End While
    End Sub

    Private Sub BeautifyTbl_Click(sender As Object, e As RibbonControlEventArgs) Handles BeautifyTblBtn.Click
        progressHandler.ProgressStart()
        RenewAllTableFormats(sender)
        progressHandler.ProgressEnd()
    End Sub
    Private Sub RenewAllTableFormats(sender As Object)
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If
        Dim tableCount As Integer = currentDoc.Tables.Count

        If Not sender Is Nothing AndAlso tableCount = 0 Then
            MessageBox.Show("文档中没有表格。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        For i As Integer = 1 To tableCount
            Dim table As Object = currentDoc.Tables(i)
            Dim percent As Integer = CInt((i / tableCount) * 100)
            progressHandler.UpdateProgress(percent, "处理第" & i & "个表格。")
            ' 检查表格是否有框线并且行和列都大于2
            If TableHasBorders(table) AndAlso table.Rows.Count > 2 AndAlso table.Columns.Count > 2 Then
                Try
                    ResetTable(table)
                    SetTableFormat(table)
                Catch
                End Try
            End If
        Next
        If Not sender Is Nothing Then
            MessageBox.Show("所有表格格式已更新。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
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
        wordApp.Selection.Cells.Borders(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth100pt

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

    Private Sub MergeTbl_Click(sender As Object, e As RibbonControlEventArgs) Handles MergeTblBtn.Click
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If
        ' 开启当前文档的修订模式
        currentDoc.TrackRevisions = False
        progressHandler.ProgressStartWaiting()
        SearchAndExecuteUnsplit()
        progressHandler.ProgressEnd()
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
            wordApp.Selection.HomeKey(Word.WdUnits.wdStory)
        End While
    End Sub

    Private Sub unsplittab(para As Word.Paragraph)
        ' 将光标移到段落的第一行
        para.Range.Select()
        wordApp.Selection.HomeKey(Word.WdUnits.wdLine)
        ' 扩展选定区域到段落的最后一行
        wordApp.Selection.MoveEnd(Word.WdUnits.wdLine, Word.WdMovementType.wdExtend + 1)

        ' 删除选定区域
        wordApp.Selection.Delete()

        ' 检查段落是否包含表格
        If para.Range.Tables.Count > 0 Then
            ' 示例操作：选择整行
            para.Range.Select()
            wordApp.Selection.Cells.Borders(Word.WdBorderType.wdBorderTop).LineWidth = Word.WdLineWidth.wdLineWidth025pt
            Try
                ResetTable(para.Range.Tables(1))
                SetTableFormat(para.Range.Tables(1))
            Catch
            End Try
        End If
    End Sub

    Private Sub SplitTbl_Click(sender As Object, e As RibbonControlEventArgs) Handles SplitTblBtn.Click
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If
        ' 切换到单页视图
        wordApp.ActiveWindow.View.Type = WdViewType.wdPrintView
        progressHandler.ProgressStartWaiting()
        SplitTable(sender)
        progressHandler.ProgressEnd()
    End Sub
    ' 这个子程序用于处理Microsoft Word中跨页的表格拆分
    Private Sub SplitTableBatch(sender As Object)
        Dim para As Paragraph
        Dim range As Range

        For Each para In currentDoc.Paragraphs
            If Not para.Style Is Nothing AndAlso para.Style.NameLocal.Contains("表标题") Then
                ' 检查当前段落后是否紧跟着一个表格
                If Not para.Range.Next(WdUnits.wdParagraph).Tables.Count = 0 Then
                    ' 将光标定位到紧跟着的表格的第一个单元格
                    range = para.Range.Next(WdUnits.wdParagraph).Tables(1).Cell(1, 1).Range
                    wordApp.Selection.SetRange(range.Start, range.End)
                    ' 调用SplitTable2过程
                    SplitTable(sender)
                End If
            End If
        Next
    End Sub

    Private Sub SplitTable(sender As Object)
        Try
            ' 循环，直到没有跨页的表格
            Do
                ' 检查当前选择是否在表格中
                If Not sender Is Nothing AndAlso Not CType(wordApp.Selection.Information(Word.WdInformation.wdWithInTable), Boolean) Then
                    MessageBox.Show("请将光标移到待拆分的表格内！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    Exit Do
                End If

                Dim selectedTable As Word.Table = wordApp.Selection.Tables(1)
                Dim startPageNumber As Integer = CType(wordApp.Selection.Information(Word.WdInformation.wdActiveEndPageNumber), Integer)
                Dim endPageNumber As Integer = CType(selectedTable.Rows(selectedTable.Rows.Count).Range.Information(Word.WdInformation.wdActiveEndPageNumber), Integer)

                ' 如果选中的表格没有跨页，则退出循环
                If Not sender Is Nothing AndAlso startPageNumber = endPageNumber Then
                    MessageBox.Show("所选表格不再跨页！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    Exit Do
                End If

                ' 选中表格
                selectedTable.Select()

                Dim selection As Word.Selection = wordApp.Selection
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
                'selection.MoveUp(Missing.Value, 1, Type.Missing)
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
        'newTable.Rows.Add(BeforeRow:=newTable.Rows(1))

        ' 复制原始表格的表头到新表格的第一行
        originalTable.Rows(1).Range.Copy()

        ' 粘贴到新表格的第一行，假设新表格已经有了一行
        newTable.Rows(1).Range.Paste()

        ' 如果需要调整边框样式，可以在这里进行
        ' 以下代码为所有新表头单元格设置底部边框样式
        Dim cell As Word.Cell
        For Each cell In newTable.Rows(1).Cells
            cell.Borders(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleSingle
            cell.Borders(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth100pt
        Next
    End Sub

    Private Sub CopyHeader2(ByRef originalTable As Word.Table, ByRef newTable As Word.Table)
        ' 在新表格的顶部插入一行
        newTable.Rows.Add(BeforeRow:=newTable.Rows(1))

        ' 复制原始表格的表头到新表格的第一行
        Dim iCol As Integer
        For iCol = 1 To originalTable.Columns.Count
            newTable.Cell(1, iCol).Range.Text = originalTable.Cell(1, iCol).Range.Text.Replace(vbCr, String.Empty)
        Next
        newTable.Rows(1).Select()
        wordApp.Selection.Cells.Borders(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth100pt
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

    Private Sub CoverChk_Click(sender As Object, e As RibbonControlEventArgs) Handles CoverChkBtn.Click
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If
        progressHandler.ProgressStart()
        ' 查找具有特定样式的段落并替换空格，同时删除末尾多余空格
        Dim uppercaseWord As String = ProcessFirstPageParagraphs(currentDoc, "标准文件_文件名称", True)
        progressHandler.UpdateProgress(50, "检查封面中文标准文件名")
        ' 查找具有特定样式的段落并格式化英文名称，同时删除末尾多余空格
        ProcessFirstPageParagraphs(currentDoc, "封面标准英文名称", False, uppercaseWord)
        progressHandler.UpdateProgress(100, "检查封面英文标准文件名")
        ' 查找具有特定样式的段落并格式化英文名称，同时删除末尾多余空格
        'Com不允许编辑：ProcessFirstPageParagraphs(currentDoc, "标准文件_正文标准名称", False)
        progressHandler.ProgressEnd()
    End Sub
    ' 处理指定样式的段落
    Private Function ProcessFirstPageParagraphs(doc As Word.Document, styleName As String, replaceSpaces As Boolean, Optional words As String = "") As String
        Dim uppercaseWord As String = ""
        For Each para As Word.Paragraph In currentDoc.Paragraphs
            If para.Style.NameLocal = "标准文件_章标题" Then Return uppercaseWord
            If para.Style.NameLocal = styleName Then
                wordApp.Selection.Start = para.Range.Start
                wordApp.Selection.End = para.Range.End - 1

                ' 删除末尾多余空格
                wordApp.Selection.Text = Trim(wordApp.Selection.Text)
                Dim regex As New Regex("[A-Z]+")
                Dim match As Match = regex.Match(wordApp.Selection.Text)
                If match.Success Then
                    uppercaseWord = match.Value ' 提取到的大写英文字符串
                End If

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
                    Return uppercaseWord
                Else
                    ' 格式化英文名称
                    wordApp.Selection.Text = LCase(wordApp.Selection.Text)
                    ' 将首个字母转为大写
                    If wordApp.Selection.Text.Length > 0 Then
                        wordApp.Selection.Text = UCase(Mid(wordApp.Selection.Text, 1, 1)) & Mid(wordApp.Selection.Text, 2)
                    End If

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
                        '替换
                    End With
                    If words.Length > 2 AndAlso wordApp.Selection.Text.ToUpper.Contains(words) Then
                        ' 保存原始选区范围
                        Dim originalRange As Word.Range = wordApp.Selection.Range

                        With wordApp.Selection.Find
                            .ClearFormatting()
                            .Text = words ' 设置查找内容
                            .Replacement.ClearFormatting()
                            .Forward = True ' 向前查找
                            .Wrap = Word.WdFindWrap.wdFindStop ' 查找到选区末尾停止
                            .Format = False ' 不使用特殊格式
                            .MatchCase = False ' 不区分大小写
                            .MatchWholeWord = True ' 全词匹配
                            .MatchWildcards = False ' 不使用通配符
                            .MatchSoundsLike = False ' 不使用发音相似查找
                            .MatchAllWordForms = False ' 不查找词的所有形式

                            ' 在选区内执行查找并替换操作
                            Do While .Execute(FindText:=words, ReplaceWith:=words,
                                           Replace:=Word.WdReplace.wdReplaceOne, Forward:=True,
                                           Wrap:=Word.WdFindWrap.wdFindStop)
                                ' 替换一次后退出
                                Return uppercaseWord
                            Loop
                        End With
                    End If
                    Return uppercaseWord
                End If
            End If
        Next
        Return uppercaseWord
    End Function

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

    Private Sub ForewordChk_Click(sender As Object, e As RibbonControlEventArgs) Handles ForewordChkBtn.Click
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If
        Dim introStarted As Boolean = False
        Dim introEnded As Boolean = False
        progressHandler.ProgressStartWaiting()
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
        progressHandler.ProgressEnd()
    End Sub

    Private Sub AbbChk_Click(sender As Object, e As RibbonControlEventArgs) Handles AbbChkBtn.Click
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If
        Dim chapterStarted As Boolean = False
        Dim chapterStartIndex As Integer = 0
        Dim paragraphsDetails As New List(Of Tuple(Of String, String))() ' 存储段落首字符和完整文本

        Dim paraCount As Integer = currentDoc.Paragraphs.Count
        Dim i As Integer = 1
        progressHandler.ProgressStartWaiting()
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
                    paraText = ReplaceWithFullWidthSpace(paraText)
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

                        progressHandler.ProgressEnd()
                        ' 退出函数
                        Exit Sub
                    End If
                End If
            Else
                i += 1
            End If
        End While
        progressHandler.ProgressEnd()
    End Sub
    Public Function ReplaceWithFullWidthSpace(input As String) As String
        ' 正则表达式分为两部分：
        ' 1. 匹配一个英文字母或数字后面紧跟的空格、冒号或破折号（假设为两个连续的减号）
        ' 2. 紧接着匹配的是第一个中文字符
        Dim regex As New Regex("([A-Za-z0-9])(\s|:|--)*([\u4e00-\u9fff])")

        ' 使用正则表达式的Replace方法，将匹配到的部分替换为英文字符或数字、一个全角空格和中文字符的组合
        Dim result As String = regex.Replace(input, Function(m) m.Groups(1).Value & ChrW(&H3000) & m.Groups(3).Value)

        Return result
    End Function

    Private Sub VarFontMdf_Click(sender As Object, e As EventArgs) Handles VarFontMdfBtn.Click
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If
        DecryptDoc(currentDoc)
        currentDoc.TrackRevisions = True

        Dim stringArray As String()
        If sender Is Nothing Then
            stringArray = New String() {"ɑ", "β", "x", "y", "z", "X", "Y", "Z", "a'", "b'", "u'", "v'", "x'", "y'", "z'", "U'", "V'", "X'", "Y'", "Z'", "u''", "v''", "x''", "y''", "z''", "U''", "V''", "X''", "Y''", "Z''"}
        Else
            Dim userInput As String = InputBox("请输入需要查找的字符或字符串:", "查找字符", "x")
            If String.IsNullOrEmpty(userInput) Then Exit Sub
            stringArray = userInput.Replace(" ", "").Split(New Char() {","c})
        End If

        Dim modifiedCount = 0
        Dim totalItems = stringArray.Length
        progressHandler.ProgressStart()

        For index As Integer = 0 To totalItems - 1
            Dim currentItem = stringArray(index)
            Dim percent As Integer = CInt((index / totalItems) * 100)

            ' 此处调用处理函数，传入currentItem作为参数
            ProcessItem(currentItem, modifiedCount)

            progressHandler.UpdateProgress(percent, "处理变量" & currentItem & "。")
        Next

        progressHandler.ProgressEnd()

        If Not sender Is Nothing Then
            If modifiedCount > 0 Then
                MessageBox.Show("发现" & modifiedCount & "个变量，已处理为斜体。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("没有发现满足条件的变量。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub

    Private Sub ProcessItem(item As String, ByRef modifiedCount As Integer)
        ' 正则表达式用于匹配非英文字符的边界
        Dim pattern As String = $"(?<=[\u4e00-\u9fff]|[，。！？,.\s])(?<![a-zA-Z])({System.Text.RegularExpressions.Regex.Escape(item)})(?![a-zA-Z])(?=[\u4e00-\u9fff]|[，。！？,.\s])"
        Dim regex As New System.Text.RegularExpressions.Regex(pattern)

        ' 遍历文档中的所有内容，查找匹配项
        Dim currentRange As Word.Range = currentDoc.Content
        currentRange.Find.ClearFormatting()
        With currentRange.Find
            .Text = item
            .MatchCase = True
            .MatchWholeWord = False  ' 修改这里，以匹配单个字符
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With

        Do While currentRange.Find.Execute
            Dim isSuperscriptOrSubscript As Boolean = currentRange.Font.Superscript <> 0 Or currentRange.Font.Subscript <> 0
            Dim isValidContext As Boolean = False
            Dim currentStart As Integer = currentRange.Start
            Dim currentEnd As Integer = currentRange.End
            Try
                ' 检查字符前后是否有英文字母，除非是上标或下标
                If Not currentDoc.Range(currentStart - 1, currentStart).Style Is Nothing AndAlso Not currentDoc.Range(currentStart - 1, currentStart).Text Is Nothing AndAlso Not currentDoc.Range(currentEnd, currentEnd + 1).Text Is Nothing AndAlso Not isSuperscriptOrSubscript Then
                    Dim isAlphabeticBefore As Boolean = currentStart > 1 AndAlso System.Text.RegularExpressions.Regex.IsMatch(currentDoc.Range(currentStart - 1, currentStart).Text, "[a-zA-Z0]")
                    Dim isAlphabeticAfter As Boolean = currentEnd < currentDoc.Content.End AndAlso System.Text.RegularExpressions.Regex.IsMatch(currentDoc.Range(currentEnd, currentEnd + 1).Text, "[a-zA-Z0-9._-]")
                    isValidContext = Not isAlphabeticAfter And Not isAlphabeticBefore
                End If
            Catch
            End Try

            If isValidContext Then
                With currentRange.Font
                    .Italic = True
                    .Name = "Times New Roman"
                End With
                modifiedCount += 1
            End If

            ' 确保不会超出文档范围
            ' 调整 Range，尝试更安全地移动到下一个位置
            Dim nextStart As Long
            If currentRange.Tables.Count > 0 Then
                ' 如果当前 Range 在表格中，尝试定位到表格之外或到下一个单元格
                nextStart = currentRange.Tables(1).Range.End + 1
            Else
                nextStart = currentRange.End + 1
            End If

            If nextStart <= currentDoc.Content.End Then
                currentRange.Start = nextStart
            Else
                Exit Do
            End If
        Loop
    End Sub
    Private Sub ListChk_Click(sender As Object, e As RibbonControlEventArgs) Handles ListChkBtn.Click
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If

        progressHandler.ProgressStart()

        ' 不是列项的名称集合
        Dim excludedStyles As New List(Of String) From {"章", "条", "标题", "附录", "图", "表", "注", "例"}
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
                    progressHandler.UpdateProgress(currentProgress, "进度：" & currentProgress & "%")

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
                        '替换末尾英文标点符号为中文标点符号，比变冒号
                        ReplaceEnglishPunctuationAtEndOfParagraph(currentDoc, para)
                        lastParagraphEndsWith = ReplaceEnglishPunctuationAtEndOfParagraph(currentDoc, para.Previous)

                        If Not isInList Then
                            ' 检查上一段落是否符合特定样式和结尾符号
                            If (lastParagraphStyle.Contains("段") OrElse lastParagraphStyle.Contains("正文")) AndAlso (lastParagraphEndsWith = "："c OrElse lastParagraphEndsWith = "。"c) Then
                                isInList = True ' 进入列项
                                listItemEndCharacter = If(lastParagraphEndsWith = "："c, "；"c, "。"c) ' 决定列项结束符
                                firstListItem = True
                            ElseIf Not missLeading Then
                                para.Range.Comments.Add(para.Range, "缺少引导语（GB/T 1.1—2020的7.5）或引导语样式不对。")
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
            progressHandler.ProgressEnd()
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
            {",", "，"},
            {"∶", "："}
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

    Private Sub SearchStd_Click(sender As Object, e As RibbonControlEventArgs) Handles SearchStdBtn.Click
        If Not IsVip Then
            MsgBox("该功能捐赠后可用。")
            Exit Sub
        End If
        Dim dialog As New BibsearchDialog()
        dialog.LicenseKey = LicenseKey
        dialog.ShowDialog()
    End Sub

    Private Sub Run_Click(sender As Object, e As RibbonControlEventArgs) Handles runBtn.Click
        If Not IsVip Then
            MsgBox("该功能捐赠后可用。")
            Exit Sub
        End If
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If

        Try
            ' 保存当前文档的状态，但不关闭它
            currentDoc.Save()
        Catch
            '如果现在是一个比较结果，也没有保存，那会到这里，就什么也不做
        End Try
        ' 保存当前文档为临时文件，作为原始副本使用
        Dim tempPathOriginal As String = Path.GetTempFileName()
        currentDoc.SaveAs2(tempPathOriginal)

        ' 解密当前文档以便进行修改
        DecryptDoc(currentDoc)

        ' 函数调用列表，每个条目都是一个无参无返回值的匿名函数（Sub）
        Dim actions As New List(Of System.Action) From {
            Sub() AcceptRev(currentDoc),
            Sub() CoverChk_Click(Nothing, e), '封面
            Sub() AcceptRev(currentDoc),
            Sub() StructureChk_Click(Nothing, e), '内容结构
            Sub() AcceptRev(currentDoc),
            Sub() ForewordChk_Click(Nothing, e),'引言
            Sub() AcceptRev(currentDoc),
            Sub() BibValid_Click(Nothing, e),'引用文件
            Sub() AcceptRev(currentDoc),
            Sub() BibRefChk_Click(Nothing, e),'引用提及
            Sub() AcceptRev(currentDoc),
            Sub() TermsChk_Click(Nothing, e),'术语
            Sub() AcceptRev(currentDoc),
            Sub() AbbChk_Click(Nothing, e),'缩略语
            Sub() AcceptRev(currentDoc),
            Sub() ListChk_Click(Nothing, e),'列项
            Sub() AcceptRev(currentDoc),
            Sub() BignumMdf_Click(Nothing, e),'千位分隔
            Sub() AcceptRev(currentDoc),
            Sub() UnitMdf_Click(Nothing, e),'量和单位
            Sub() AcceptRev(currentDoc),
            Sub() VarFontMdf_Click(Nothing, e),'变量字体
            Sub() AcceptRev(currentDoc),
            Sub() MergeTbl_Click(Nothing, e),'批量合并表格
            Sub() AcceptRev(currentDoc),
            Sub() BeautifyTbl_Click(Nothing, e),'批量美化表格
            Sub() AcceptRev(currentDoc),
            Sub() SplitTbl_Click(Nothing, e),'批量拆分表格
            Sub() AcceptRev(currentDoc),
            Sub() ApplyStyle_Click(Nothing, e), '应用样式
            Sub() AcceptRev(currentDoc)
        }

        For Each func As System.Action In actions
            Try
                ' 尝试执行当前函数
                func.Invoke()
            Catch ex As Exception
                ' 捕获到异常后的处理逻辑
                Console.WriteLine($"Error in {func.Method.Name}: {ex.Message}")
                ' 这里可以记录日志、重试或者其他自定义错误处理
            End Try
        Next

        Try
            ' 保存修改后的当前文档为另一个临时文件
            Dim tempPathModified As String = Path.GetTempFileName()
            currentDoc.SaveAs2(tempPathModified)

            ' 重新打开原始文档的副本和修改后的副本进行比较
            Dim originalDoc As Word.Document = wordApp.Documents.Open(tempPathOriginal)
            Dim modifiedDoc As Word.Document = wordApp.Documents.Open(tempPathModified)
            DecryptDoc(originalDoc)
            DecryptDoc(modifiedDoc)

            ' 使用CompareDocuments方法比较文档，生成比较结果作为新文档
            Dim comparedDocument As Word.Document = wordApp.CompareDocuments(OriginalDocument:=originalDoc, RevisedDocument:=modifiedDoc, Destination:=Word.WdCompareDestination.wdCompareDestinationNew, Granularity:=Word.WdGranularity.wdGranularityCharLevel, CompareFormatting:=False, CompareHeaders:=True, CompareFootnotes:=True, CompareTextboxes:=True, CompareFields:=True, CompareComments:=False, CompareMoves:=True, IgnoreAllComparisonWarnings:=True)

            ' 关闭临时文档
            originalDoc.Close(False)
            modifiedDoc.Close(False)

            ' 清理临时文件
            File.Delete(tempPathOriginal)
            File.Delete(tempPathModified)

            'currentDoc = comparedDocument
            ' 加密当前文档
            'EncryptDoc(currentDoc)

            ' 比较结果文档保持打开状态
            ' 注意：不需要再次打开currentDoc，因为它已经是打开状态
        Catch ex As Exception
            progressHandler.ProgressEnd()
            MsgBox("出现错误: " & ex.Message)
        End Try
    End Sub

    Private Sub ApplyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles ApplyStyleBtn.Click
        If wordApp.Documents.Count > 0 Then
            currentDoc = wordApp.ActiveDocument
        Else
            MsgBox("没有打开的文件。")
            Exit Sub
        End If

        ' 应用模板中的样式到当前文档
        currentDoc.AttachedTemplate = StylePath
        currentDoc.UpdateStyles()
    End Sub

    Public Sub DecryptDoc(doc As Document)
        'doc = Me.wordApp.ActiveDocument
        ' 检查文档是否受保护
        If Not doc Is Nothing AndAlso doc.ProtectionType <> WdProtectionType.wdNoProtection Then
            ' 如果文档受到保护，尝试解除保护
            ' 如果文档是用密码保护的，需要提供密码作为参数
            Try
                doc.Unprotect(Password:="haizi") ' 如果没有密码，可以省略这个参数或传递空字符串
            Catch
            End Try
        End If
    End Sub

    Public Sub EncryptDoc(doc As Document)
        'doc = Me.wordApp.ActiveDocument
        ' 应用完模板样式后，根据需要重新保护文档
        ' 使用适当的保护类型，例如 wdAllowOnlyReading, wdAllowOnlyFormFields 等
        ' 如果之前解除了带密码的保护，这里也需要用同样的密码重新保护
        If Not doc Is Nothing AndAlso doc.ProtectionType = WdProtectionType.wdNoProtection Then
            Try
                doc.Protect(Type:=WdProtectionType.wdAllowOnlyFormFields, NoReset:=True, Password:="haizi") ' 根据实际情况选择保护类型和是否使用密码
            Catch
            End Try
        End If
    End Sub

    Private Sub AcceptRev(doc As Document)
        ' 接受当前文档中的所有修订
        For Each revision As Word.Revision In doc.Revisions
            revision.Accept()
        Next
    End Sub

    Private Sub Donate_Click(sender As Object, e As RibbonControlEventArgs) Handles DonateBtn.Click
        Dim donateForm As New DonateForm()
        donateForm.ShowDialog() ' 以模态方式显示窗体
    End Sub

    Private Sub Setting_Click(sender As Object, e As RibbonControlEventArgs) Handles SettingBtn.Click
        Dim settingsForm As New SettingForm(Me)
        settingsForm.ShowDialog() ' 或者使用 settingsForm.Show() 根据需要
    End Sub
    Private Sub InsertCatalog()
        Dim flag As Boolean = False
        Dim num_location As Integer = 0
        For i As Integer = 2 To 4
            If wordApp.ActiveDocument.Bookmarks.Exists("BookMark" & i) Then
                flag = True
                num_location = i
                Exit For
            End If
        Next
        If flag Then
            If wordApp.ActiveDocument.Bookmarks.Exists("BookMark1") Then
                Dim activeDocument As Document = wordApp.ActiveDocument
                Dim bookmarks As Bookmarks = wordApp.ActiveDocument.Bookmarks
                Dim Index As Object = "BookMark1"
                Dim Start As Object = bookmarks(Index).Start
                Dim bookmarks2 As Bookmarks = wordApp.ActiveDocument.Bookmarks
                Index = "BookMark1"
                Dim [End] As Object = bookmarks2(Index).Start
                activeDocument.Range(Start, [End]).Select()
            End If
        Else
            MessageBox.Show("目次定位标签被删除，无法插入目次！", "错误")
        End If
    End Sub
    Public Sub LoadSettings()
        Dim assemblyPath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
        Dim filePath As String = Path.Combine(assemblyPath, IniPath)
        If File.Exists(filePath) Then
            Dim lines As String() = File.ReadAllLines(filePath)
            For Each line In lines
                Dim parts As String() = line.Split("="c)
                If parts.Length = 2 Then
                    Select Case parts(0).Trim().ToLower()
                        Case "licensekey"
                            LicenseKey = parts(1).Trim()
                        Case "llm"
                            Llm = parts(1).Trim()
                        Case "llmkey"
                            LlmKey = parts(1).Trim()
                    End Select
                End If
            Next

            ' 异步执行，不阻塞UI线程
            Threading.Tasks.Task.Run(Async Function()
                                         Me.IsVip = Await ValidLicenseKeyAsync(LicenseKey)
                                     End Function)
        End If
    End Sub
    Friend Shared Async Function ValidLicenseKeyAsync(licenseKey As String) As Threading.Tasks.Task(Of Boolean)
        ' 忽略SSL证书验证（生产环境中应处理证书验证）
        ServicePointManager.ServerCertificateValidationCallback = Function(s, certificate, chain, sslPolicyErrors) True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

        Dim success As Boolean = False
        Dim url As String = $"https://api.relaton.top:4567/validkey?key={licenseKey}"
        Using client As New HttpClient()
            Try
                Dim response As HttpResponseMessage = Await client.GetAsync(url)
                response.EnsureSuccessStatusCode()
                Dim responseBody As String = Await response.Content.ReadAsStringAsync()

                Dim data As Dictionary(Of String, Boolean) = JsonConvert.DeserializeObject(Of Dictionary(Of String, Boolean))(responseBody)

                success = If(data.ContainsKey("valid"), data("valid"), False)
            Catch e As Exception
                Console.WriteLine($"An error occurred: {e.Message}")
                success = False
            End Try
        End Using
        Return success
    End Function

    Private Sub AIwriting_Click(sender As Object, e As RibbonControlEventArgs) Handles AIwriting.Click
        MsgBox("该功能正在开发，请捐赠。")
        Exit Sub
    End Sub
End Class
