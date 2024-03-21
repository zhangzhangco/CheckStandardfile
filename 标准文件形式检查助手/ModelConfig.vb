Public Class ModelConfig
    Public Property ApiKey As String
    Public Property ModelId As String
    Public Property Parameters As Dictionary(Of String, Object)

    ' 修改构造函数，使其能够使用全局变量
    Public Sub New()
        Me.ApiKey = Ribbon1.LlmKey
        Me.ModelId = Ribbon1.Llm
        Parameters = New Dictionary(Of String, Object)()
    End Sub
    Public Sub New(apiSelection As String, apiKey As String)
        Me.ModelId = apiSelection
        Me.ApiKey = apiKey
    End Sub
End Class
