Imports System.Windows.Forms

Public Class ProgressForm
    Inherits Form
    Friend WithEvents ProgressLabel1 As Label
    Friend WithEvents ProgressBar1 As ProgressBar
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub UpdateProgress(value As Integer)
        ' 更新进度条的值
        ProgressBar1.Value = value
    End Sub
    Public Sub UpdateMessage(value As String)
        ' 更新进度条的值
        ProgressLabel1.Text = value
    End Sub

    Public Sub InitializeComponent()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.ProgressLabel1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(-2, 12)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(412, 40)
        Me.ProgressBar1.TabIndex = 0
        '
        'ProgressLabel1
        '
        Me.ProgressLabel1.AutoEllipsis = True
        Me.ProgressLabel1.AutoSize = True
        Me.ProgressLabel1.Location = New System.Drawing.Point(22, 74)
        Me.ProgressLabel1.Name = "ProgressLabel1"
        Me.ProgressLabel1.Size = New System.Drawing.Size(0, 24)
        Me.ProgressLabel1.TabIndex = 2
        Me.ProgressLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ProgressForm
        '
        Me.ClientSize = New System.Drawing.Size(408, 157)
        Me.Controls.Add(Me.ProgressLabel1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ProgressForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
End Class