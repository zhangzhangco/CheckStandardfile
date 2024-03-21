'在与大模型的交互中，文本的预处理和后处理也重要。
'TextProcessor 可以用来清洗和准备发送到模型的数据（比如去除多余空格、特殊字符等），以及解析和格式化模型返回的结果，使其适用于后续操作。
'例如，如果模型返回的是 JSON 字符串，TextProcessor 可以负责解析这些数据并提取出有用的信息。
Public Class TextProcessor
    Public Shared Function PreProcess(text As String) As String
        ' 实现文本的预处理逻辑
        Return text
    End Function

    Public Shared Function PostProcess(result As String) As String
        ' 实现对模型返回结果的后处理逻辑
        Return result
    End Function
End Class
