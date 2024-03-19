<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DonateForm
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DonateForm))
        Me.wechatPaymentPictureBox = New System.Windows.Forms.PictureBox()
        Me.wechatQRCodePictureBox = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.wechatPaymentPictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.wechatQRCodePictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'wechatPaymentPictureBox
        '
        Me.wechatPaymentPictureBox.Image = CType(resources.GetObject("wechatPaymentPictureBox.Image"), System.Drawing.Image)
        Me.wechatPaymentPictureBox.Location = New System.Drawing.Point(57, 123)
        Me.wechatPaymentPictureBox.Name = "wechatPaymentPictureBox"
        Me.wechatPaymentPictureBox.Size = New System.Drawing.Size(384, 436)
        Me.wechatPaymentPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.wechatPaymentPictureBox.TabIndex = 0
        Me.wechatPaymentPictureBox.TabStop = False
        '
        'wechatQRCodePictureBox
        '
        Me.wechatQRCodePictureBox.Image = CType(resources.GetObject("wechatQRCodePictureBox.Image"), System.Drawing.Image)
        Me.wechatQRCodePictureBox.Location = New System.Drawing.Point(537, 123)
        Me.wechatQRCodePictureBox.Name = "wechatQRCodePictureBox"
        Me.wechatQRCodePictureBox.Size = New System.Drawing.Size(456, 436)
        Me.wechatQRCodePictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.wechatQRCodePictureBox.TabIndex = 1
        Me.wechatQRCodePictureBox.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(119, 54)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(236, 28)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "第一步：微信支付"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(647, 54)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(236, 28)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "第二步：添加好友"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(63, 606)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(706, 48)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "💡 每一份捐赠，都是对创新和卓越的投资。 💡" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "您的支持会激励我不断改进完善此软件，助您更高效地完成工作。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button1.Location = New System.Drawing.Point(807, 606)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(126, 47)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "谢谢！"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'DonateForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1048, 720)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.wechatQRCodePictureBox)
        Me.Controls.Add(Me.wechatPaymentPictureBox)
        Me.Name = "DonateForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "捐赠"
        CType(Me.wechatPaymentPictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.wechatQRCodePictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents wechatPaymentPictureBox As Windows.Forms.PictureBox
    Friend WithEvents wechatQRCodePictureBox As Windows.Forms.PictureBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Button1 As Windows.Forms.Button
End Class
