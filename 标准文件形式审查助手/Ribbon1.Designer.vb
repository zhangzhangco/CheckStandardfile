Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
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

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Button11 = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button12 = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Button13 = Me.Factory.CreateRibbonButton
        Me.Button15 = Me.Factory.CreateRibbonButton
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Button14 = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Button10 = Me.Factory.CreateRibbonButton
        Me.Button8 = Me.Factory.CreateRibbonButton
        Me.Button9 = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.searchStd = Me.Factory.CreateRibbonButton
        Me.Button7 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Label = "形式审查助手"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button11)
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Items.Add(Me.Button12)
        Me.Group1.Items.Add(Me.Button2)
        Me.Group1.Items.Add(Me.Button4)
        Me.Group1.Items.Add(Me.Button3)
        Me.Group1.Items.Add(Me.Button13)
        Me.Group1.Items.Add(Me.Button15)
        Me.Group1.Items.Add(Me.Button5)
        Me.Group1.Items.Add(Me.Button6)
        Me.Group1.Items.Add(Me.Button14)
        Me.Group1.Label = "形式检查项"
        Me.Group1.Name = "Group1"
        '
        'Button11
        '
        Me.Button11.Image = CType(resources.GetObject("Button11.Image"), System.Drawing.Image)
        Me.Button11.Label = "检查封面"
        Me.Button11.Name = "Button11"
        Me.Button11.ShowImage = True
        Me.Button11.SuperTip = "检查中文和英文的文件名格式"
        '
        'Button1
        '
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Label = "内容结构"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        Me.Button1.SuperTip = "1.删除多余空行;2.批注出悬置段;3.批注冗余标题;4.批注缺失标题"
        '
        'Button12
        '
        Me.Button12.Image = CType(resources.GetObject("Button12.Image"), System.Drawing.Image)
        Me.Button12.Label = "检查引言"
        Me.Button12.Name = "Button12"
        Me.Button12.ShowImage = True
        Me.Button12.SuperTip = "检查时候含有要求性表述"
        '
        'Button2
        '
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Label = "引用文件"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        Me.Button2.SuperTip = "1.调整顺序;2.修正标点符号"
        '
        'Button4
        '
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.Label = "引用提及"
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        Me.Button4.SuperTip = "检查提及的文件是否存在于规范性引用文件和参考文献"
        '
        'Button3
        '
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.Label = "检查术语"
        Me.Button3.Name = "Button3"
        Me.Button3.ShowImage = True
        Me.Button3.SuperTip = "1.检查术语在文中是否出现两次以上;2.修复set2020的Bug(错误的回车和样式格式)"
        '
        'Button13
        '
        Me.Button13.Image = CType(resources.GetObject("Button13.Image"), System.Drawing.Image)
        Me.Button13.Label = "查缩略语"
        Me.Button13.Name = "Button13"
        Me.Button13.ShowImage = True
        Me.Button13.SuperTip = "1.缩略语排序;2.是否在文中出现"
        '
        'Button15
        '
        Me.Button15.Image = CType(resources.GetObject("Button15.Image"), System.Drawing.Image)
        Me.Button15.Label = "检查列项"
        Me.Button15.Name = "Button15"
        Me.Button15.ShowImage = True
        '
        'Button5
        '
        Me.Button5.Image = CType(resources.GetObject("Button5.Image"), System.Drawing.Image)
        Me.Button5.Label = "千位分隔"
        Me.Button5.Name = "Button5"
        Me.Button5.ShowImage = True
        Me.Button5.SuperTip = "有理数整数部分达5位,使用千位分隔符(,)"
        '
        'Button6
        '
        Me.Button6.Image = CType(resources.GetObject("Button6.Image"), System.Drawing.Image)
        Me.Button6.Label = "量与单位"
        Me.Button6.Name = "Button6"
        Me.Button6.ShowImage = True
        Me.Button6.SuperTip = "统一使用阿拉伯数字和英文单位符号,中间适当使用半角空格"
        '
        'Button14
        '
        Me.Button14.Image = CType(resources.GetObject("Button14.Image"), System.Drawing.Image)
        Me.Button14.Label = "变量斜体"
        Me.Button14.Name = "Button14"
        Me.Button14.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.Button10)
        Me.Group3.Items.Add(Me.Button8)
        Me.Group3.Items.Add(Me.Button9)
        Me.Group3.Label = "表格处理"
        Me.Group3.Name = "Group3"
        '
        'Button10
        '
        Me.Button10.Image = CType(resources.GetObject("Button10.Image"), System.Drawing.Image)
        Me.Button10.Label = "连续合并"
        Me.Button10.Name = "Button10"
        Me.Button10.ShowImage = True
        '
        'Button8
        '
        Me.Button8.Image = CType(resources.GetObject("Button8.Image"), System.Drawing.Image)
        Me.Button8.Label = "批量美化"
        Me.Button8.Name = "Button8"
        Me.Button8.ShowImage = True
        Me.Button8.SuperTip = "1.连续表格美化;2.一字线填充空缺;3.删除末尾的句号"
        '
        'Button9
        '
        Me.Button9.Image = CType(resources.GetObject("Button9.Image"), System.Drawing.Image)
        Me.Button9.Label = "连续拆分"
        Me.Button9.Name = "Button9"
        Me.Button9.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.searchStd)
        Me.Group2.Items.Add(Me.Button7)
        Me.Group2.Name = "Group2"
        '
        'searchStd
        '
        Me.searchStd.Image = CType(resources.GetObject("searchStd.Image"), System.Drawing.Image)
        Me.searchStd.Label = "标准查询"
        Me.searchStd.Name = "searchStd"
        Me.searchStd.ShowImage = True
        Me.searchStd.SuperTip = "查询国内外标准状态"
        '
        'Button7
        '
        Me.Button7.Image = CType(resources.GetObject("Button7.Image"), System.Drawing.Image)
        Me.Button7.Label = "关于"
        Me.Button7.Name = "Button7"
        Me.Button7.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button7 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button8 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button9 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button10 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button11 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button12 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button13 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button14 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button15 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents searchStd As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
