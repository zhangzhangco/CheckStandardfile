Public Class ModelInterfaceFactory
    Public Shared Function CreateModelInterface() As IModelInterface
        Select Case Ribbon.Llm
            Case "OpenAI"
                Return New OpenAIModelInterface()
            Case "智普AI" ' 假设这是另一种模型的标识
                Return New ZhiPuModelInterface() ' 假设这是智普模型的接口实现
            Case Else
                Throw New Exception("未知的模型类型")
        End Select
    End Function
End Class
