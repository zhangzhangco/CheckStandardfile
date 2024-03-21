'可以用来验证输入数据的安全性，防止注入攻击，或限制与大模型交互时可执行的操作范围。
'例如， 确保传递给大模型的文本不包含可能触发不当行为的敏感信息。
Public Class SecurityManager
    Public Shared Function IsRequestValid(text As String) As Boolean
        ' 实现安全检查的逻辑，确保请求是合法的
        Return True
    End Function
End Class
