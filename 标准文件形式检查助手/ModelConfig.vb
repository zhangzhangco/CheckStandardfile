Public Class ModelConfig
    Public Property ApiKey As String
    Public Property ApiToken As String
    Public Property ModelId As String
    Public Property Parameters As Dictionary(Of String, Object)

    ' 修改构造函数，使其能够使用全局变量
    Public Sub New()
        Me.ApiKey = Ribbon.LlmKey
        Me.ModelId = Ribbon.Llm
        Me.ApiToken = Ribbon.LlmToken
        Parameters = New Dictionary(Of String, Object)()
    End Sub
    Public Sub New(apiSelection As String, apiKey As String, apiToken As String)
        Me.ModelId = apiSelection
        Me.ApiKey = apiKey
        Me.ApiToken = apiToken
    End Sub
End Class
