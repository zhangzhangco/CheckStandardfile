
Public Interface IModelInterface
    Function SendToLLM(prompt As String, apiToken As String) As String
    Function TestApiKey(config As ModelConfig) As Boolean
    Function CheckRequirements(text As String, config As ModelConfig) As Boolean ' 用于分析文本
    Function OptimizeSentence(sentence As String, config As ModelConfig) As String ' 用于优化句子
    Function WriteTerm(text As String, config As ModelConfig) As String
    Function WriteClause(text As String, config As ModelConfig) As String
    ' 根据需要定义其他方法
End Interface
