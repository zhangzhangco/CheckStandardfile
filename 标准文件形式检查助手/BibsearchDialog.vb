Imports System.Net
Imports System.Net.Http
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Clipboard = System.Windows.Forms.Clipboard
Imports MessageBox = System.Windows.Forms.MessageBox

Public Class BibsearchDialog
    Inherits Form

    ' 使用静态HttpClient实例以提高效率和资源复用
    Public Shared ReadOnly HttpClientInstance As New HttpClient()
    Public Property LicenseKey As String

    Friend WithEvents searchCancel As Button
    Friend WithEvents searchOk As Button
    Friend WithEvents rb_guonei As RadioButton
    Friend WithEvents rb_guoji As RadioButton
    Friend WithEvents tb_stdCode As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel1 As Panel

    Public Sub New()
        InitializeComponent()
    End Sub
    Public Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.searchCancel = New System.Windows.Forms.Button()
        Me.searchOk = New System.Windows.Forms.Button()
        Me.rb_guonei = New System.Windows.Forms.RadioButton()
        Me.rb_guoji = New System.Windows.Forms.RadioButton()
        Me.tb_stdCode = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.searchCancel)
        Me.Panel1.Controls.Add(Me.searchOk)
        Me.Panel1.Location = New System.Drawing.Point(2, 185)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(365, 92)
        Me.Panel1.TabIndex = 0
        '
        'searchCancel
        '
        Me.searchCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.searchCancel.Location = New System.Drawing.Point(217, 22)
        Me.searchCancel.Name = "searchCancel"
        Me.searchCancel.Size = New System.Drawing.Size(92, 46)
        Me.searchCancel.TabIndex = 1
        Me.searchCancel.Text = "取消"
        Me.searchCancel.UseVisualStyleBackColor = True
        '
        'searchOk
        '
        Me.searchOk.Location = New System.Drawing.Point(54, 22)
        Me.searchOk.Name = "searchOk"
        Me.searchOk.Size = New System.Drawing.Size(92, 46)
        Me.searchOk.TabIndex = 0
        Me.searchOk.Text = "确定"
        Me.searchOk.UseVisualStyleBackColor = True
        '
        'rb_guonei
        '
        Me.rb_guonei.AutoSize = True
        Me.rb_guonei.Checked = True
        Me.rb_guonei.Location = New System.Drawing.Point(3, 3)
        Me.rb_guonei.Name = "rb_guonei"
        Me.rb_guonei.Size = New System.Drawing.Size(89, 28)
        Me.rb_guonei.TabIndex = 1
        Me.rb_guonei.TabStop = True
        Me.rb_guonei.Text = "国内"
        Me.rb_guonei.UseVisualStyleBackColor = True
        '
        'rb_guoji
        '
        Me.rb_guoji.AutoSize = True
        Me.rb_guoji.Location = New System.Drawing.Point(121, 3)
        Me.rb_guoji.Name = "rb_guoji"
        Me.rb_guoji.Size = New System.Drawing.Size(89, 28)
        Me.rb_guoji.TabIndex = 2
        Me.rb_guoji.Text = "国际"
        Me.rb_guoji.UseVisualStyleBackColor = True
        '
        'tb_stdCode
        '
        Me.tb_stdCode.Location = New System.Drawing.Point(56, 128)
        Me.tb_stdCode.Name = "tb_stdCode"
        Me.tb_stdCode.Size = New System.Drawing.Size(265, 35)
        Me.tb_stdCode.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(52, 82)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(202, 24)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "请输入标准编号："
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.rb_guonei)
        Me.Panel2.Controls.Add(Me.rb_guoji)
        Me.Panel2.Location = New System.Drawing.Point(56, 25)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(243, 54)
        Me.Panel2.TabIndex = 5
        '
        'BibsearchDialog
        '
        Me.AcceptButton = Me.searchOk
        Me.CancelButton = Me.searchCancel
        Me.ClientSize = New System.Drawing.Size(377, 281)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tb_stdCode)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "BibsearchDialog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "标准查询"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Public Sub searchOk_Click(control As Object, e As EventArgs) Handles searchOk.Click
        If tb_stdCode.Text <> "" Then
            Search(tb_stdCode.Text.ToUpper(), rb_guonei.Checked, HttpClientInstance)
        End If
    End Sub
    Public Sub Search(ByVal text As String, ByVal isDomestic As Boolean, HttpClientInstance As HttpClient)
        ' 忽略SSL证书验证（生产环境中应处理证书验证）
        ServicePointManager.ServerCertificateValidationCallback = Function(control, certificate, chain, sslPolicyErrors) True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        ' 或者，为了兼容未来的协议版本，可以这样设置：
        ' ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13

        ' 根据用户选择处理code值
        Dim code As String = If(isDomestic, $"{text.Replace("—", "-")}", text)
        Dim out As String = String.Empty
        Dim pattern As String = "^([A-Z0-9]+(/[TZ])?)\s(\d+(\.\d+)?)[-—]?(\d{4})?$"
        If isDomestic Then
            If Not Regex.IsMatch(text, pattern) Then
                ' 使用正则表达式验证文本
                MessageBox.Show("中文编号错误。示例：GB/T 1.1-2020"）
            Else
                '由自己的过程处理国内标准
                out = SearchCnGovStd(code, HttpClientInstance)
            End If
        Else
            '查询国外
            out = SearchInterStd(code, HttpClientInstance, LicenseKey)
        End If

        If String.IsNullOrEmpty(out) Then
            MessageBox.Show("标准不存在或已废止")
        Else
            MessageBox.Show($"文档有效: {out}")
            Clipboard.SetText(out)
            MessageBox.Show("信息已复制到剪贴板。")
        End If
    End Sub
End Class
