Imports System.Drawing
Imports System.Net
Imports System.Net.Http
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles
Imports Microsoft.Office.Interop.Word
Imports Newtonsoft.Json.Linq
Imports Clipboard = System.Windows.Forms.Clipboard
Imports MessageBox = System.Windows.Forms.MessageBox

Public Class BibsearchDialog
    Inherits Form

    Friend WithEvents searchCancel As Button
    Friend WithEvents searchOk As Button
    Friend WithEvents rb_guonei As RadioButton
    Friend WithEvents rb_guoji As RadioButton
    Friend WithEvents tb_stdCode As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel1 As Panel

    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.searchCancel = New System.Windows.Forms.Button()
        Me.searchOk = New System.Windows.Forms.Button()
        Me.rb_guonei = New System.Windows.Forms.RadioButton()
        Me.rb_guoji = New System.Windows.Forms.RadioButton()
        Me.tb_stdCode = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.searchCancel)
        Me.Panel1.Controls.Add(Me.searchOk)
        Me.Panel1.Location = New System.Drawing.Point(2, 129)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(270, 54)
        Me.Panel1.TabIndex = 0
        '
        'searchCancel
        '
        Me.searchCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.searchCancel.Location = New System.Drawing.Point(153, 12)
        Me.searchCancel.Name = "searchCancel"
        Me.searchCancel.Size = New System.Drawing.Size(75, 23)
        Me.searchCancel.TabIndex = 1
        Me.searchCancel.Text = "取消"
        Me.searchCancel.UseVisualStyleBackColor = True
        '
        'searchOk
        '
        Me.searchOk.Location = New System.Drawing.Point(32, 12)
        Me.searchOk.Name = "searchOk"
        Me.searchOk.Size = New System.Drawing.Size(75, 23)
        Me.searchOk.TabIndex = 0
        Me.searchOk.Text = "确定"
        Me.searchOk.UseVisualStyleBackColor = True
        '
        'rb_guonei
        '
        Me.rb_guonei.AutoSize = True
        Me.rb_guonei.Checked = True
        Me.rb_guonei.Location = New System.Drawing.Point(3, 3)
        Me.rb_guonei.Name = "rb_guonei"
        Me.rb_guonei.Size = New System.Drawing.Size(47, 16)
        Me.rb_guonei.TabIndex = 1
        Me.rb_guonei.TabStop = True
        Me.rb_guonei.Text = "国内"
        Me.rb_guonei.UseVisualStyleBackColor = True
        '
        'rb_guoji
        '
        Me.rb_guoji.AutoSize = True
        Me.rb_guoji.Location = New System.Drawing.Point(121, 3)
        Me.rb_guoji.Name = "rb_guoji"
        Me.rb_guoji.Size = New System.Drawing.Size(47, 16)
        Me.rb_guoji.TabIndex = 2
        Me.rb_guoji.Text = "国际"
        Me.rb_guoji.UseVisualStyleBackColor = True
        '
        'tb_stdCode
        '
        Me.tb_stdCode.Location = New System.Drawing.Point(32, 88)
        Me.tb_stdCode.Name = "tb_stdCode"
        Me.tb_stdCode.Size = New System.Drawing.Size(202, 21)
        Me.tb_stdCode.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(32, 66)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(101, 12)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "请输入标准编号："
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.rb_guonei)
        Me.Panel2.Controls.Add(Me.rb_guoji)
        Me.Panel2.Location = New System.Drawing.Point(34, 32)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(200, 28)
        Me.Panel2.TabIndex = 5
        '
        'BibsearchDialog
        '
        Me.AcceptButton = Me.searchOk
        Me.CancelButton = Me.searchCancel
        Me.ClientSize = New System.Drawing.Size(269, 183)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tb_stdCode)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "BibsearchDialog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "标准查询"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private Sub searchOk_Click(sender As Object, e As EventArgs) Handles searchOk.Click
        If tb_stdCode.Text <> "" Then
            Search(tb_stdCode.Text.ToUpper(), rb_guonei.Checked)
        End If
    End Sub
    Private Sub Search(ByVal text As String, ByVal isDomestic As Boolean)
        ' 忽略SSL证书验证（生产环境中应处理证书验证）
        ServicePointManager.ServerCertificateValidationCallback = Function(sender, certificate, chain, sslPolicyErrors) True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        ' 或者，为了兼容未来的协议版本，可以这样设置：
        ' ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13

        ' 根据用户选择处理code值
        Dim code As String = If(isDomestic, $"{text.Replace("—", "-")}", text)

        Dim pattern As String = "^([A-Z0-9]+(/[TZ])?)\s(\d+(\.\d+)?)[-—]?(\d{4})?$"
        If isDomestic Then
            If Not Regex.IsMatch(text, pattern) Then
                ' 使用正则表达式验证文本
                MessageBox.Show("中文编号错误。示例：GB/T 1.1-2020"）
            Else
                '由自己的过程处理国内标准
                Dim output As String = SearchCnGovStd(code)

                If String.IsNullOrEmpty(output) Then
                    MessageBox.Show("标准不存在或已废止")
                Else
                    MessageBox.Show($"文档有效: {output}")
                    Clipboard.SetText(output)
                    MessageBox.Show("信息已复制到剪贴板。")
                End If
            End If
        Else
            ' 异步发送请求并获取JSON响应
            Dim jsonString As String = FetchJson($"https://39.96.136.172:4567/fetch?code={code}&retries=2")
            If Not String.IsNullOrEmpty(jsonString) Then
                AnalyzeAndShowMessageBox(jsonString)
            Else
                MessageBox.Show("未查到")
            End If
        End If
    End Sub
    Private Function FetchJson(url As String) As String
        Try
            ' 构造HttpRequestMessage以允许添加请求头部
            Dim requestMessage As New HttpRequestMessage(HttpMethod.Get, url)

            ' 添加必要的请求头部信息
            requestMessage.Headers.Accept.ParseAdd("text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7")
            requestMessage.Headers.AcceptLanguage.ParseAdd("zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6")
            requestMessage.Headers.CacheControl = New Net.Http.Headers.CacheControlHeaderValue With {.MaxAge = TimeSpan.FromSeconds(0)}
            requestMessage.Headers.Add("Upgrade-Insecure-Requests", "1")

            ' 使用HttpClient实例发送构造好的请求
            Dim response As HttpResponseMessage = Ribbon1.HttpClientInstance.SendAsync(requestMessage).GetAwaiter().GetResult()
            If response.IsSuccessStatusCode Then
                ' 确保同步调用
                Return response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            Else
                'MessageBox.Show("请求失败: " & response.ReasonPhrase)
                Return Nothing
            End If
        Catch ex As HttpRequestException
            MessageBox.Show($"HTTP请求异常: {ex.Message}")
        Catch ex As TaskCanceledException
            MessageBox.Show("请求超时")
        Catch ex As Exception
            MessageBox.Show($"未处理的异常: {ex.Message}")
        End Try
        Return Nothing
    End Function

    Private Sub AnalyzeAndShowMessageBox(jsonString As String)
        ' 解析JSON响应
        Dim jsonObj As JObject = JObject.Parse(jsonString)
        Dim stageValue As String = If(jsonObj.SelectToken("docstatus.stage.value") IsNot Nothing, jsonObj.SelectToken("docstatus.stage.value").ToString(), "")
        Dim isDocumentActive As Boolean = stageValue = "" OrElse stageValue = "60" OrElse stageValue = "90" OrElse stageValue.ToLower() = "activated" OrElse stageValue.ToUpper() = "PUBLISHED"

        If isDocumentActive Then
            ' 提取docid和title
            Dim docId As String = jsonObj.SelectTokens("$.docid[?(@.primary == true)].id").FirstOrDefault()?.ToString()

            ' 尝试获取中文标题，如果没有中文标题，尝试获取英文标题
            Dim titleContent As String = GetPreferredTitle(jsonObj, "zh")
            Dim originalTitleContent As String = titleContent
            If String.IsNullOrEmpty(titleContent) Then
                titleContent = GetPreferredTitle(jsonObj, "en")
                ' 如果是英文标题，尝试翻译
                If Not String.IsNullOrEmpty(titleContent) Then
                    Dim translatedTitle As String = TranslateText(titleContent)
                    titleContent = $"{translatedTitle}（{titleContent}）"
                End If
            Else
                docId = docId.Replace("-", "—")
                titleContent = titleContent.Replace("  ", "　").Replace(" ", "　")
            End If

            Dim resultString As String = If(docId IsNot Nothing, $"{docId}　", "") & titleContent.Replace("免责声明：本平台标准资料仅供参考，使用标准请以正式出版的标准版本为准。", "")

            ' 显示消息框并复制到剪贴板
            If Not String.IsNullOrEmpty(resultString.Trim()) Then
                MessageBox.Show($"文档有效: {resultString}")
                Clipboard.SetText(resultString)
                MessageBox.Show("信息已复制到剪贴板。")
            End If
        Else
            MessageBox.Show("未查到或文件废止/被替换")
        End If
    End Sub
    ' 根据语言获取首选标题
    Private Function GetPreferredTitle(jsonObj As JObject, languageCode As String) As String
        For Each title In jsonObj("title")
            Dim lang = title("language").FirstOrDefault()?.ToString()
            If lang IsNot Nothing AndAlso lang.StartsWith(languageCode) Then
                Return title("content").ToString()
            End If
        Next
        Return String.Empty
    End Function
    Private Function TranslateText(ByVal text As String) As String
        ' 微软翻译API调用逻辑
        ' 请确保替换以下URL和headers中的subscriptionKey和region等值为您的实际值
        Dim endpoint As String = "https://api.cognitive.microsofttranslator.com/"
        Dim route As String = "/translate?api-version=3.0&from=en&to=zh"
        Dim subscriptionKey As String = "09f169299acb45a9bfb8e99dc5d6ba0a"
        Dim region As String = "eastasia"

        Using client As New HttpClient()
            Dim request As New HttpRequestMessage()
            request.Method = HttpMethod.Post
            request.RequestUri = New Uri(endpoint & route)
            request.Content = New StringContent("[{""Text"":""" & text & """}]", Encoding.UTF8, "application/json")
            request.Headers.Add("Ocp-Apim-Subscription-Key", subscriptionKey)
            request.Headers.Add("Ocp-Apim-Subscription-Region", region)

            Dim responseTask = client.SendAsync(request)
            responseTask.Wait() ' 阻塞调用线程直到任务完成
            Dim response = responseTask.Result ' 获取结果

            If response.IsSuccessStatusCode Then
                Dim readTask = response.Content.ReadAsStringAsync()
                readTask.Wait() ' 同样阻塞等待
                Dim responseBody = readTask.Result ' 获取结果

                ' 解析响应体以获取翻译结果
                Dim jsonResponse = JArray.Parse(responseBody)
                Dim translation = jsonResponse(0)("translations")(0)("text").ToString()

                Return translation
            Else
                Throw New HttpRequestException($"请求翻译服务失败: {response.ReasonPhrase}")
            End If
        End Using
    End Function
End Class
