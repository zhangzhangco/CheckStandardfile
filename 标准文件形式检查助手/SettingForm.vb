Imports System.Drawing
Imports System.Net
Imports System.Net.Http
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles
Imports System.IO
Imports Application = System.Windows.Forms.Application
Imports System.Reflection
Imports System.Diagnostics
Public Class SettingForm
    Inherits Form

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents OK As Button
    Friend WithEvents Cancel As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label1 As Label
    Friend WithEvents licenseKeyTB As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents LlmKeyTB As TextBox
    Friend WithEvents LlmCB As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents TstLicenseBtn As Button
    Friend WithEvents TstLlmBtn As Button

    ' 使用静态HttpClient实例以提高效率和资源复用
    Public Shared ReadOnly HttpClientInstance As New HttpClient()
    Friend WithEvents update As Button
    Private _rb As Ribbon
    Private llmApiToken As String

    Public Sub New(rb As Ribbon)
        InitializeComponent()
        _rb = rb
    End Sub
    Public Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TstLicenseBtn = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.licenseKeyTB = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.TstLlmBtn = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LlmKeyTB = New System.Windows.Forms.TextBox()
        Me.LlmCB = New System.Windows.Forms.ComboBox()
        Me.OK = New System.Windows.Forms.Button()
        Me.Cancel = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.update = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TstLicenseBtn)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.licenseKeyTB)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(51, 59)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(692, 168)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "尊享授权"
        '
        'TstLicenseBtn
        '
        Me.TstLicenseBtn.Location = New System.Drawing.Point(586, 50)
        Me.TstLicenseBtn.Name = "TstLicenseBtn"
        Me.TstLicenseBtn.Size = New System.Drawing.Size(75, 36)
        Me.TstLicenseBtn.TabIndex = 3
        Me.TstLicenseBtn.Text = "测试"
        Me.TstLicenseBtn.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(57, 108)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(394, 24)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "捐赠后加微信 HelloLLM2035 获取。"
        '
        'licenseKeyTB
        '
        Me.licenseKeyTB.Location = New System.Drawing.Point(136, 51)
        Me.licenseKeyTB.Name = "licenseKeyTB"
        Me.licenseKeyTB.Size = New System.Drawing.Size(422, 35)
        Me.licenseKeyTB.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(57, 54)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Key："
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.TstLlmBtn)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.LlmKeyTB)
        Me.GroupBox2.Controls.Add(Me.LlmCB)
        Me.GroupBox2.Location = New System.Drawing.Point(51, 261)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(692, 182)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "大模型"
        '
        'TstLlmBtn
        '
        Me.TstLlmBtn.Location = New System.Drawing.Point(586, 120)
        Me.TstLlmBtn.Name = "TstLlmBtn"
        Me.TstLlmBtn.Size = New System.Drawing.Size(75, 35)
        Me.TstLlmBtn.TabIndex = 4
        Me.TstLlmBtn.Text = "测试"
        Me.TstLlmBtn.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 123)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(118, 24)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "API Key："
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(45, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 24)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "名称："
        '
        'LlmKeyTB
        '
        Me.LlmKeyTB.Location = New System.Drawing.Point(136, 120)
        Me.LlmKeyTB.Name = "LlmKeyTB"
        Me.LlmKeyTB.Size = New System.Drawing.Size(422, 35)
        Me.LlmKeyTB.TabIndex = 2
        '
        'LlmCB
        '
        Me.LlmCB.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.LlmCB.FormattingEnabled = True
        Me.LlmCB.Items.AddRange(New Object() {"智普AI", "OpenAI"})
        Me.LlmCB.Location = New System.Drawing.Point(136, 53)
        Me.LlmCB.Name = "LlmCB"
        Me.LlmCB.Size = New System.Drawing.Size(207, 32)
        Me.LlmCB.TabIndex = 1
        Me.LlmCB.Tag = ""
        '
        'OK
        '
        Me.OK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OK.Location = New System.Drawing.Point(435, 499)
        Me.OK.Name = "OK"
        Me.OK.Size = New System.Drawing.Size(133, 39)
        Me.OK.TabIndex = 2
        Me.OK.Text = "确定"
        Me.OK.UseVisualStyleBackColor = True
        '
        'Cancel
        '
        Me.Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel.Location = New System.Drawing.Point(610, 499)
        Me.Cancel.Name = "Cancel"
        Me.Cancel.Size = New System.Drawing.Size(133, 39)
        Me.Cancel.TabIndex = 3
        Me.Cancel.Text = "取消"
        Me.Cancel.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Location = New System.Drawing.Point(13, 13)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(759, 456)
        Me.Panel1.TabIndex = 4
        '
        'update
        '
        Me.update.Location = New System.Drawing.Point(37, 499)
        Me.update.Name = "update"
        Me.update.Size = New System.Drawing.Size(112, 39)
        Me.update.TabIndex = 5
        Me.update.Text = "更新"
        Me.update.UseVisualStyleBackColor = True
        '
        'SettingForm
        '
        Me.AcceptButton = Me.OK
        Me.CancelButton = Me.Cancel
        Me.ClientSize = New System.Drawing.Size(784, 567)
        Me.Controls.Add(Me.update)
        Me.Controls.Add(Me.Cancel)
        Me.Controls.Add(Me.OK)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "SettingForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "设置"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Public Sub SettingForm_Load(control As Object, e As EventArgs) Handles MyBase.Load
        LoadSettings()
    End Sub
    Public Sub OK_Click(control As Object, e As EventArgs) Handles OK.Click
        SaveSettings()
    End Sub
    Public Sub SaveSettings()
        ' 获取当前执行的DLL的目录
        Dim assemblyPath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
        Dim filePath As String = Path.Combine(assemblyPath, Ribbon.IniPath)

        ' 接下来，保存设置到setting.ini文件
        Using writer As New StreamWriter(filePath, False)
            writer.WriteLine("licensekey=" & licenseKeyTB.Text)
            writer.WriteLine("llm=" & LlmCB.Text)
            writer.WriteLine("llmkey=" & LlmKeyTB.Text)
            writer.WriteLine("llmtoken=" & Me.llmApiToken)
        End Using
        _rb.LoadSettings()
    End Sub

    Public Sub LoadSettings()
        ' 获取当前执行的DLL的目录
        Dim assemblyPath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
        Dim filePath As String = Path.Combine(assemblyPath, Ribbon.IniPath)

        If File.Exists(filePath) Then
            ' 文件存在时，加载设置
            Dim lines As String() = File.ReadAllLines(filePath)
            For Each line In lines
                Dim parts As String() = line.Split("="c)
                If parts.Length = 2 Then
                    Select Case parts(0).Trim().ToLower()
                        Case "licensekey"
                            licenseKeyTB.Text = parts(1).Trim()
                        Case "llm"
                            LlmCB.SelectedItem = parts(1).Trim()
                        Case "llmkey"
                            LlmKeyTB.Text = parts(1).Trim()
                    End Select
                End If
            Next
        End If
    End Sub

    Public Sub TstLicenseBtn_Click(control As Object, e As EventArgs) Handles TstLicenseBtn.Click
        System.Threading.Tasks.Task.Run(Async Function()
                                            If Not Await Ribbon.ValidLicenseKeyAsync(licenseKeyTB.Text) Then
                                                Forms.MessageBox.Show("测试失败。", "测试结果", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                            Else
                                                Forms.MessageBox.Show("测试成功。", "测试结果", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                            End If
                                        End Function)
    End Sub

    Public Sub update_Click(control As Object, e As EventArgs) Handles update.Click
        Dim process1 As Process = Process.Start(Ribbon.UpdaterPath, "/checknow")
    End Sub

    Private Function TstLlmBtn_Click(control As Object, e As EventArgs) Handles TstLlmBtn.Click
        Dim apiSelection As String = LlmCB.SelectedItem.ToString()
        Dim apiKey As String = LlmKeyTB.Text
        Dim apkToken As String = Nothing
        Dim modelInterface As IModelInterface
        If apiSelection = "OpenAI" Then
            modelInterface = New OpenAIModelInterface()
        Else
            modelInterface = New ZhiPuModelInterface()
            apkToken = ZhiPuModelInterface.GenerateToken(apiKey, 864000)
        End If

        Dim config As New ModelConfig(apiSelection, apiKey, apkToken)
        Dim isSuccess As Boolean = modelInterface.TestApiKey(config)
        If isSuccess Then
            Me.llmApiToken = apkToken
            Forms.MessageBox.Show("API key有效。", "测试成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Forms.MessageBox.Show("API key无效，请检查。", "测试失败", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Function
End Class
