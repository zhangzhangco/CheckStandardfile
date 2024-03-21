Imports Newtonsoft.Json
Imports System.Net.Http
Imports System.Threading.Tasks

Public Class OpenAIModelInterface
    Implements IModelInterface

    Private ReadOnly httpClient As HttpClient

    Public Sub New()
        httpClient = New HttpClient()
    End Sub
    Public Function TestApiKey(config As ModelConfig) As Boolean Implements IModelInterface.TestApiKey
        Try
            ' 发送一个简单的请求来测试API key
            Dim prompt As String = "The following is a test: Hello, world!"
            SendToOpenAI(prompt, config.ApiKey)
            Return True ' 如果没有异常，则认为key有效
        Catch ex As Exception
            Return False ' 如果有异常，则认为key无效
        End Try
    End Function
    Public Function CheckRequirements(text As String, config As ModelConfig) As Boolean Implements IModelInterface.CheckRequirements
        Try
            ' 构建专用于检查要求性条款的prompt
            Dim prompt As String = $"检查以下文本是否包含要求性条款（使用了能愿动词：应、不应、应该、不应该、只准许、不准许、必须、不可等），仅回答我是或否： '{text}'"
            Dim response As String = SendToOpenAI(prompt, config.ApiKey)
            ' 解析response并返回结果
            Return InterpretResponse(response)
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function OptimizeSentenceAsync(sentence As String, config As ModelConfig) As String Implements IModelInterface.OptimizeSentence
        Throw New NotImplementedException()
    End Function

    Private Function SendToOpenAI(prompt As String, apiKey As String) As String
        ' 初始化HttpClient
        If httpClient.DefaultRequestHeaders.Authorization Is Nothing Then
            httpClient.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey)
        End If

        ' 构造请求体
        Dim requestBody = New With {
        .model = "gpt-3.5-turbo",
        .messages = New List(Of Object) From {
            New With {.role = "user", .content = prompt}
        },
        .temperature = 0.7
    }

        ' 将请求体对象转换为JSON字符串
        Dim jsonContent = JsonConvert.SerializeObject(requestBody)
        Dim content = New StringContent(jsonContent, Encoding.UTF8, "application/json")

        ' 发送POST请求到OpenAI的chat/completions端点
        Dim response = httpClient.PostAsync("https://api.openai.com/v1/chat/completions", content).Result

        ' 确保请求成功
        response.EnsureSuccessStatusCode()

        ' 返回响应内容
        Return response.Content.ReadAsStringAsync().Result
    End Function

    ' 这是一个示例函数，需要根据OpenAI返回的实际内容进行实现
    Private Function InterpretResponse(response As String) As Boolean
        ' 解析JSON响应
        Dim jsonResponse = Newtonsoft.Json.Linq.JObject.Parse(response)
        ' 获取第一个choice的content
        Dim content As String = jsonResponse("choices")(0)("message")("content").ToString()
        If content.Contains("是") Then
            Return True
        End If
        Return False
    End Function
End Class
