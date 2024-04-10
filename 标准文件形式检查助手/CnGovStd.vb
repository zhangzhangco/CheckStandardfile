Imports System.IO ' 用于处理字符串中的正则表达式
Imports System.Net
Imports System.Net.Http
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq

Module CnGovStd
    Sub Main()
        Dim HttpClientInstance As New HttpClient()
        ' 定义需要查询的标准代码数组
        Dim codes As String() = New String() {"GB/T 1.1", "GB/T 1.1-2009", "JB/T 6166", "JB/T 6167-2008", "GB/T 1.1-2020", "GB/T 1.1—2020", "DY/T 7-2023", "GB/T 1.1", "JB/T 234.1"}

        For Each code In codes
            Dim output As String = SearchCnGovStd(code, HttpClientInstance)

            If String.IsNullOrEmpty(output) Then
                Console.WriteLine("标准不存在或已废止")
            Else
                Console.WriteLine(output)
            End If
        Next

        Console.ReadLine()
    End Sub

    Public Function SearchCnGovStd(code As String, HttpClientInstance As HttpClient) As String
        ' 忽略SSL证书验证（生产环境中应处理证书验证）
        ServicePointManager.ServerCertificateValidationCallback = Function(control, certificate, chain, sslPolicyErrors) True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Dim result As String = String.Empty
        code = code.Replace("—", "-")
        Dim originalCode As String = code
        Dim processedCode As String = ProcessCode(code)

        ' 根据标准代码的前缀决定查询的网址
        If processedCode.StartsWith("GB") Then
            ' 查询国家标准
            result = SendHttpRequest("https://std.samr.gov.cn/gb/search/gbQueryPage?searchText=" & processedCode & "&ics=&state=&ISSUE_DATE=&sortOrder=asc&pageSize=15&pageNumber=1&_=1708353434376", "")
        Else
            ' 查询行业标准
            result = SendHttpRequest("https://hbba.sacinfo.org.cn/stdQueryList", "current=1&size=15&key=" & processedCode & "&ministry=&industry=&pubdate=&date=&status=")
        End If

        ' 处理返回的JSON数据
        Dim output As String = ProcessJson(result, originalCode)
        Return output
    End Function

    ' 移除标准代码中的年份
    Function ProcessCode(code As String) As String
        Return Regex.Replace(code, "-\d{4}|—\d{4}", "")
    End Function

    ' 发送HTTP请求，可以处理GET和POST请求
    Function SendHttpRequest(url As String, postData As String) As String
        Dim request As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)

        If Not String.IsNullOrEmpty(postData) Then
            request.Method = "POST"
            request.ContentType = "application/x-www-form-urlencoded"
            Using streamWriter As New StreamWriter(request.GetRequestStream())
                streamWriter.Write(postData)
            End Using
        Else
            request.Method = "GET"
        End If

        Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Using streamReader As New StreamReader(response.GetResponseStream())
                Return streamReader.ReadToEnd()
            End Using
        End Using
    End Function


    ' 处理返回的JSON数据，寻找并返回现行的标准
    Function ProcessJson(jsonString As String, originalCode As String) As String
        Dim json As JObject = JObject.Parse(jsonString)
        ' 尝试获取国家标准或行业标准的数组
        Dim standardsArray As JArray = If(json("rows") IsNot Nothing, json("rows"), If(json("records") IsNot Nothing, json("records"), New JArray()))

        ' 根据原始代码决定是否需要处理年份信息
        Dim needsYearRemoval As Boolean = Not (originalCode.Contains("-") OrElse originalCode.Contains("—"))

        For Each item In standardsArray
            Dim status As String = If(item("STATE")?.ToString(), item("status")?.ToString())
            If status = "现行" Then
                Dim code As String = If(item("C_STD_CODE")?.ToString(), item("code")?.ToString())
                Dim name As String = If(item("C_C_NAME")?.ToString(), item("chName")?.ToString())
                '修复网站提供数据的错误
                code = Regex.Replace(code, "<.*?>", String.Empty)
                name = Regex.Replace(name, "<.*?>", String.Empty)
                ' 对于不需要年份的情况，返回不带年份的标准代码和名称
                If needsYearRemoval Then code = ProcessCode(code)
                If Not String.IsNullOrEmpty(code) AndAlso Not String.IsNullOrEmpty(name) Then
                    Return StandardDocument.FormatedFileName($"{code.Replace("-", "—")}　{name.Replace("  ", "　").Replace(" ", "　")}")
                End If
            End If
        Next

        ' 如果没有找到符合条件的记录，返回空字符串
        Return String.Empty
    End Function
End Module
