Imports Microsoft.VisualBasic
Imports System.Windows.Forms

Public Class ProgressForm
    Inherits Form

    Public Property ProgressBar As System.Windows.Forms.ProgressBar
    Public Property ProgressLabel As System.Windows.Forms.Label

    Public Sub New()
        ' 设置窗口的基本属性
        Me.FormBorderStyle = FormBorderStyle.None ' 无标题栏
        Me.TopMost = True ' 窗口置顶
        Me.StartPosition = FormStartPosition.CenterParent ' 窗口位置
        Me.Size = New Size(300, 100) ' 窗口大小

        ' 初始化进度条
        ProgressBar = New System.Windows.Forms.ProgressBar()
        ProgressBar.Dock = DockStyle.Top
        ProgressBar.Minimum = 0
        ProgressBar.Maximum = 100
        Me.Controls.Add(ProgressBar)

        ' 初始化进度信息标签
        ProgressLabel = New System.Windows.Forms.Label()
        ProgressLabel.Dock = DockStyle.Fill
        ProgressLabel.TextAlign = ContentAlignment.MiddleCenter
        Me.Controls.Add(ProgressLabel)

        ' 设置窗体为无法通过鼠标点击穿透
        Me.Enabled = False
    End Sub
End Class
