Imports System.Net
Imports System.Net.Http
Imports Newtonsoft.Json.Linq
Imports System.Threading.Tasks
Imports System.Windows.Forms

Module InternationalStd
    Public Function SearchInterStd(code As String, HttpClientInstance As HttpClient, licenseKey As String) As String
        ' 忽略SSL证书验证（生产环境中应处理证书验证）
        ServicePointManager.ServerCertificateValidationCallback = Function(control, certificate, chain, sslPolicyErrors) True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

        ' 异步发送请求并获取JSON响应
        Dim jsonString As String = FetchJson($"https://api.relaton.top:4567/fetch?code={code.Replace("—", "-")}&retries=2" & "&key=" & licenseKey)
        If Not String.IsNullOrEmpty(jsonString) Then
            Return AnalyzeAndShowMessageBox(jsonString)
        Else
            Return String.Empty
        End If
    End Function
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
            Dim response As HttpResponseMessage = Ribbon.HttpClientInstance.SendAsync(requestMessage).GetAwaiter().GetResult()
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

    Private Function AnalyzeAndShowMessageBox(jsonString As String) As String
        ' 解析JSON响应
        Dim jsonObj As JObject = JObject.Parse(jsonString)
        If (jsonObj.SelectToken("error")) Then
            Return String.Empty
        End If
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
                    Dim translatedTitle As String = TranslateText(titleContent).Replace(" ", "")
                    If Not String.IsNullOrEmpty(translatedTitle) Then
                        titleContent = $"{translatedTitle}（{titleContent}）"
                    End If
                Else
                    docId = docId.Replace("-", "—")
                    titleContent = titleContent.Replace("  ", "　").Replace(" ", "　")
                End If
            End If
            Return If(docId IsNot Nothing, $"{docId}　", "") & titleContent.Replace("免责声明：本平台标准资料仅供参考，使用标准请以正式出版的标准版本为准。", "")

        Else
            Return String.Empty
        End If
    End Function
    ' 根据语言获取首选标题
    Private Function GetPreferredTitle(jsonObj As JObject, languageCode As String) As String
        ' 遍历 "title" 数组中的每个对象
        For Each title In jsonObj("title")
            ' 获取当前对象的 "type" 值
            Dim type As String = title("type").ToString()

            ' 检查 "type" 是否为 "main"
            If type = "main" Then
                ' 检查 "language" 数组中是否包含指定的语言代码
                Dim languages = title("language").ToObject(Of List(Of String))()
                If languages.Contains(languageCode) Then
                    ' 如果找到，则返回对应的 "content"
                    Return title("content").ToString()
                End If
            End If
        Next

        ' 如果没有找到匹配的项，返回空字符串
        Return String.Empty
    End Function

    Private Function TranslateText(ByVal text As String) As String
        ' 微软翻译API调用逻辑
        ' 请确保替换以下URL和headers中的subscriptionKey和region等值为您的实际值
        Dim endpoint As String = "https://api.cognitive.microsofttranslator.com/"
        Dim route As String = "/translate?api-version=3.0&from=en&to=zh"
        Dim subscriptionKey As String = "09f169299acb45a9bfb8e99dc5d6ba0a"
        Dim region As String = "eastasia"
        Try

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

                    Return translation.Replace(" - ", "　")
                Else
                    Throw New HttpRequestException($"请求翻译服务失败: {response.ReasonPhrase}")
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show($"未处理的异常: {ex.Message}")
            Return "网络异常请重试"
        End Try
    End Function
End Module
