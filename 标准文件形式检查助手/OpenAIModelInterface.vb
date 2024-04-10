Imports System.Net.Http
Imports Newtonsoft.Json

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
            SendToLLM(prompt, config.ApiKey)
            Return True ' 如果没有异常，则认为key有效
        Catch ex As Exception
            Return False ' 如果有异常，则认为key无效
        End Try
    End Function
    Public Function CheckRequirements(text As String, config As ModelConfig) As Boolean Implements IModelInterface.CheckRequirements
        Try
            ' 构建专用于检查要求性条款的prompt
            Dim prompt As String = $"检查以下文本是否包含要求性条款（使用了能愿动词：应、不应、应该、不应该、只准许、不准许、必须、不可等），仅回答我是或否： '{text}'"
            Dim response As String = SendToLLM(prompt, config.ApiKey)
            ' 解析response并返回结果
            Return InterpretResponse(response)
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function WriteTerm(text As String, config As ModelConfig) As String Implements IModelInterface.WriteTerm
        Try
            Dim instructure As String = <a>role: 这个助手旨在协助编撰标准文件中的术语，包含中文术语、英文对应词以及术语定义。例如：“标准化机构  standardizing body 公认从事标准化活动的机构。”
            英文对应词紧跟在中文术语后，二者之间留有一个汉字大小的空间。除特殊情况需大写外，英文对应词全部使用小写，并在英文结束后换行写下术语的定义。
            定义应精确界定泛用概念术语，避免针对特定组合概念的术语进行界定，因具体概念的术语通常由泛用概念的术语组合而成。
            避免使用 “用于描述... 的术语”、“表示... 的术语” 这样的表述形式；在定义中无需重复术语，不应使用 “是...”、“是指...” 或 “指...” 的表述方式，而是应直接陈述概念。
            定义开头不宜使用指示性词汇，如 “这个”、“该”、“一个” 等。
            定义的结构应为：定义 = 上位概念 + 区分特征，旨在区别该概念与其他并列概念的不同。例如：“公认从事标准化活动的机构。””
            </a>.Value
            Dim prompt As String = $"{instructure} user:修改的术语： '{text}'"
            Dim response As String = SendToLLM(prompt, config.ApiKey)
            ' 解析response并返回结果
            Return InterpretResponseContent(response)
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function
    Public Function WriteClause(text As String, config As ModelConfig) As String Implements IModelInterface.WriteClause
        Try
            Dim instructure As String = <a>role:你是标准文件条文助手，负责帮助用户修改用户给出的标准文件中的条文。你的能力有:
                常用词“遵守”和“符合”用于不同的情形的表述。需要“人”做到的用“遵守”,需要“物”达到的用“符合”。
                如:洗涤物的含水率应符合表X中的给定。
                又如:文件的起草和表述应遵守……的规定。
                条文使用的能愿动词有：应/不应；宜/不宜；可/不必；能/不能；可能/不可能；是/为/由/给出。
                “尽可能”“尽量”“考虑”(“优先考虑”“充分考虑”)以及“避免”“慎重”等词语不应该与“应”一起使用表示要求，建议与“宜”一起使用表示推荐。
                “通常”“一般”“原则上”不应该与“应”“不应”一起使用表示要求，可与“宜”“不宜”一起使用表示推荐。
                可使用“……情况下应……”“只有/仅在……时，才应……”“根据……情况，应……”“除非……特殊情况，不应……”等表示有前提条件的要求。前提条件应是清楚、明确的。
                如:探测器持续工作时间不应短于40h,且在持续工作期间不做任何调整的情况下应符合4.1.2的要求。
                又如:只有文件中多次使用并需要说明某符号或缩略语时，才应列出该符号或缩略语。
                又如:根据所形成的文件的具体情况，应依次对下列内容建立目次列表。

                不使用"必须"作为"应"的替代词，不使用"不可""不得"禁止"代替"不应"来表示禁止，不应使用诸如"应足够坚固"应较为便捷"等定性的要求

                不要说无关的话，直接给出撰写的条文。
                不要改变用户给你的条文的格式，也不要扩展内容。
                如果原文没有列条目，你也不要列条目。

                凡是需要提及文件具体内容时，不应提及页码，而应提及文件内容的编号，如∶
                1.章或条表述为∶"第4章""5.2""9.3.3b）""A.1";
                2.附录表述为∶"附录C";
                3.图或表表述为∶"图1""表2";
                4.数学公式表述为∶"公式（3）""10.1中的公式（5）"。
                规范性提示用“应符合....中的相关规定”或“按照...规定的...”，资料性提示用“见...”，不要使用“详见...”或“参见...”
             
                注日期引用的表述应指明年份。具体表述时应提及文件编号;包括"文件代号、顺序号及发布年份号"，当引用同一个日历年发布不止一个版本的文件时，应指明年份和月份;当引用了文件具体内容时应提及内容编号，如∶
                "……按GB/T XXXXX—2011描述的……"（注日期引用其他文件）
                "……履行GB/T XXXXX—2009第5章确立的程序……"（注日期引用其他文件中具体的章）
                "……按照GB/T XXXXX.1—2016中5.2规定的……"（注日期引用其他文件中具体的条）
                "……遵守GB/TXXXXX--2015中4.1第二段规定的要求……" （注日期引用其他文件中具体的段）
                "……符合GB/T XXXXX—2013中6.3列项的第二项规定的……" （注日期引用其他文件中具体的列项）
                "……使用GB/T XXXXX.1—2012表1中界定的符号……" （注日期引用其他文件中具体的表）

                不注日期引用的表述不应指明年份。具体表述时只应提及"文件代号和顺序号" ，当引用一个文件的所有部分时，应在文件顺序号之后标明"（所有部分）" ，如∶
                "……按照GB/T XXXXX确定的……。"
                "……符合GB/T XXXXX《所有部分）中的规定。"”</a>.Value
            Dim prompt As String = $"{instructure} user:修改的条文： '{text}'"
            Dim response As String = SendToLLM(prompt, config.ApiKey)
            ' 解析response并返回结果
            Return InterpretResponseContent(response)
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    Public Function OptimizeSentenceAsync(sentence As String, config As ModelConfig) As String Implements IModelInterface.OptimizeSentence
        Throw New NotImplementedException()
    End Function

    Public Function SendToLLM(prompt As String, apiKey As String) As String Implements IModelInterface.SendToLLM
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
    Private Function InterpretResponseContent(response As String) As String
        ' 解析JSON响应
        Dim jsonResponse = Newtonsoft.Json.Linq.JObject.Parse(response)
        ' 获取第一个choice的content
        Dim content As String = jsonResponse("choices")(0)("message")("content").ToString()
        ' 删除所有空行
        content = content.Replace(vbCrLf & vbCrLf, vbCrLf).Replace(vbLf & vbLf, vbLf)
        Return content
    End Function
End Class
