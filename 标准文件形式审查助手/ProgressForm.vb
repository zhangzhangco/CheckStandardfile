Imports System.Drawing
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles

Public Class ProgressForm
    Inherits Form

    Public Property ProgressBar As System.Windows.Forms.ProgressBar
    Public Property ProgressLabel As System.Windows.Forms.Label

    Public Sub New()
        ' 设置窗口的基本属性
        Me.FormBorderStyle = FormBorderStyle.None ' 无标题栏
        Me.TopMost = True ' 窗口置顶
        'Me.StartPosition = FormStartPosition.CenterParent ' 窗口位置
        Me.Size = New Drawing.Size(300, 100) ' 窗口大小
        ' 在构造函数中添加以下代码来启用双缓冲
        Me.DoubleBuffered = True
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True)
        Me.UpdateStyles()

        ' 初始化进度条
        ProgressBar = New System.Windows.Forms.ProgressBar()
        ProgressBar.Dock = DockStyle.Top
        ProgressBar.Minimum = 0
        ProgressBar.Maximum = 100
        Me.Controls.Add(ProgressBar)

        ' 初始化进度信息标签
        ProgressLabel = New System.Windows.Forms.Label()
        ProgressLabel.Dock = DockStyle.Fill
        ProgressLabel.TextAlign = Drawing.ContentAlignment.MiddleCenter
        Me.Controls.Add(ProgressLabel)

        ' 设置窗体为无法通过鼠标点击穿透
        Me.Enabled = False
    End Sub

    Public Sub UpdateProgress(value As Integer)
        ' 更新进度条的值
        ProgressBar.Value = value
    End Sub
    Public Sub UpdateMessage(value As String)
        ' 更新进度条的值
        ProgressLabel.Text = value
    End Sub
End Class