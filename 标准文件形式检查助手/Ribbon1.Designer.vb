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
        Me.CoverChkBtn = Me.Factory.CreateRibbonButton
        Me.StructureChkBtn = Me.Factory.CreateRibbonButton
        Me.ForewordChkBtn = Me.Factory.CreateRibbonButton
        Me.BibValidBtn = Me.Factory.CreateRibbonButton
        Me.BibRefChkBtn = Me.Factory.CreateRibbonButton
        Me.TermsChkBtn = Me.Factory.CreateRibbonButton
        Me.AbbChkBtn = Me.Factory.CreateRibbonButton
        Me.ListChkBtn = Me.Factory.CreateRibbonButton
        Me.BignumMdfBtn = Me.Factory.CreateRibbonButton
        Me.UnitMdfBtn = Me.Factory.CreateRibbonButton
        Me.VarFontMdfBtn = Me.Factory.CreateRibbonButton
        Me.ApplyStyleBtn = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.MergeTblBtn = Me.Factory.CreateRibbonButton
        Me.BeautifyTblBtn = Me.Factory.CreateRibbonButton
        Me.SplitTblBtn = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.SearchStdBtn = Me.Factory.CreateRibbonButton
        Me.runBtn = Me.Factory.CreateRibbonButton
        Me.AIwriting = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.DonateBtn = Me.Factory.CreateRibbonButton
        Me.SettingBtn = Me.Factory.CreateRibbonButton
        Me.AboutBtn = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Label = "形式检查助手"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.CoverChkBtn)
        Me.Group1.Items.Add(Me.StructureChkBtn)
        Me.Group1.Items.Add(Me.ForewordChkBtn)
        Me.Group1.Items.Add(Me.BibValidBtn)
        Me.Group1.Items.Add(Me.BibRefChkBtn)
        Me.Group1.Items.Add(Me.TermsChkBtn)
        Me.Group1.Items.Add(Me.AbbChkBtn)
        Me.Group1.Items.Add(Me.ListChkBtn)
        Me.Group1.Items.Add(Me.BignumMdfBtn)
        Me.Group1.Items.Add(Me.UnitMdfBtn)
        Me.Group1.Items.Add(Me.VarFontMdfBtn)
        Me.Group1.Items.Add(Me.ApplyStyleBtn)
        Me.Group1.Label = "形式检查项"
        Me.Group1.Name = "Group1"
        '
        'CoverChkBtn
        '
        Me.CoverChkBtn.Image = CType(resources.GetObject("CoverChkBtn.Image"), System.Drawing.Image)
        Me.CoverChkBtn.Label = "检查封面"
        Me.CoverChkBtn.Name = "CoverChkBtn"
        Me.CoverChkBtn.ShowImage = True
        Me.CoverChkBtn.SuperTip = "检查中文和英文的文件名格式"
        '
        'StructureChkBtn
        '
        Me.StructureChkBtn.Image = CType(resources.GetObject("StructureChkBtn.Image"), System.Drawing.Image)
        Me.StructureChkBtn.Label = "内容结构"
        Me.StructureChkBtn.Name = "StructureChkBtn"
        Me.StructureChkBtn.ShowImage = True
        Me.StructureChkBtn.SuperTip = "1.删除多余空行;2.批注出悬置段;3.批注冗余标题;4.批注缺失标题"
        '
        'ForewordChkBtn
        '
        Me.ForewordChkBtn.Image = CType(resources.GetObject("ForewordChkBtn.Image"), System.Drawing.Image)
        Me.ForewordChkBtn.Label = "检查引言"
        Me.ForewordChkBtn.Name = "ForewordChkBtn"
        Me.ForewordChkBtn.ShowImage = True
        Me.ForewordChkBtn.SuperTip = "检查时候含有要求性表述"
        '
        'BibValidBtn
        '
        Me.BibValidBtn.Image = CType(resources.GetObject("BibValidBtn.Image"), System.Drawing.Image)
        Me.BibValidBtn.Label = "引用文件"
        Me.BibValidBtn.Name = "BibValidBtn"
        Me.BibValidBtn.ShowImage = True
        Me.BibValidBtn.SuperTip = "1.调整顺序;2.修正标点符号;3.引文有效性"
        '
        'BibRefChkBtn
        '
        Me.BibRefChkBtn.Image = CType(resources.GetObject("BibRefChkBtn.Image"), System.Drawing.Image)
        Me.BibRefChkBtn.Label = "引用提及"
        Me.BibRefChkBtn.Name = "BibRefChkBtn"
        Me.BibRefChkBtn.ShowImage = True
        Me.BibRefChkBtn.SuperTip = "检查提及的文件是否存在于规范性引用文件和参考文献"
        '
        'TermsChkBtn
        '
        Me.TermsChkBtn.Image = CType(resources.GetObject("TermsChkBtn.Image"), System.Drawing.Image)
        Me.TermsChkBtn.Label = "检查术语"
        Me.TermsChkBtn.Name = "TermsChkBtn"
        Me.TermsChkBtn.ShowImage = True
        Me.TermsChkBtn.SuperTip = "1.检查术语在文中是否出现两次以上;2.修复set2020的Bug(错误的回车和样式格式)"
        '
        'AbbChkBtn
        '
        Me.AbbChkBtn.Image = CType(resources.GetObject("AbbChkBtn.Image"), System.Drawing.Image)
        Me.AbbChkBtn.Label = "查缩略语"
        Me.AbbChkBtn.Name = "AbbChkBtn"
        Me.AbbChkBtn.ShowImage = True
        Me.AbbChkBtn.SuperTip = "1.缩略语排序;2.是否在文中出现"
        '
        'ListChkBtn
        '
        Me.ListChkBtn.Image = CType(resources.GetObject("ListChkBtn.Image"), System.Drawing.Image)
        Me.ListChkBtn.Label = "检查列项"
        Me.ListChkBtn.Name = "ListChkBtn"
        Me.ListChkBtn.ShowImage = True
        '
        'BignumMdfBtn
        '
        Me.BignumMdfBtn.Image = CType(resources.GetObject("BignumMdfBtn.Image"), System.Drawing.Image)
        Me.BignumMdfBtn.Label = "千位分隔"
        Me.BignumMdfBtn.Name = "BignumMdfBtn"
        Me.BignumMdfBtn.ShowImage = True
        Me.BignumMdfBtn.SuperTip = "有理数整数部分达5位,使用千位分隔符(,)"
        '
        'UnitMdfBtn
        '
        Me.UnitMdfBtn.Image = CType(resources.GetObject("UnitMdfBtn.Image"), System.Drawing.Image)
        Me.UnitMdfBtn.Label = "量与单位"
        Me.UnitMdfBtn.Name = "UnitMdfBtn"
        Me.UnitMdfBtn.ShowImage = True
        Me.UnitMdfBtn.SuperTip = "统一使用阿拉伯数字和英文单位符号,中间适当使用半角空格"
        '
        'VarFontMdfBtn
        '
        Me.VarFontMdfBtn.Image = CType(resources.GetObject("VarFontMdfBtn.Image"), System.Drawing.Image)
        Me.VarFontMdfBtn.Label = "变量斜体"
        Me.VarFontMdfBtn.Name = "VarFontMdfBtn"
        Me.VarFontMdfBtn.ShowImage = True
        '
        'ApplyStyleBtn
        '
        Me.ApplyStyleBtn.Image = CType(resources.GetObject("ApplyStyleBtn.Image"), System.Drawing.Image)
        Me.ApplyStyleBtn.Label = "间距修复"
        Me.ApplyStyleBtn.Name = "ApplyStyleBtn"
        Me.ApplyStyleBtn.ScreenTip = "修复错误的中文与西文、数字的间距"
        Me.ApplyStyleBtn.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.MergeTblBtn)
        Me.Group3.Items.Add(Me.BeautifyTblBtn)
        Me.Group3.Items.Add(Me.SplitTblBtn)
        Me.Group3.Label = "表格处理"
        Me.Group3.Name = "Group3"
        '
        'MergeTblBtn
        '
        Me.MergeTblBtn.Image = CType(resources.GetObject("MergeTblBtn.Image"), System.Drawing.Image)
        Me.MergeTblBtn.Label = "连续合并"
        Me.MergeTblBtn.Name = "MergeTblBtn"
        Me.MergeTblBtn.ShowImage = True
        '
        'BeautifyTblBtn
        '
        Me.BeautifyTblBtn.Image = CType(resources.GetObject("BeautifyTblBtn.Image"), System.Drawing.Image)
        Me.BeautifyTblBtn.Label = "批量美化"
        Me.BeautifyTblBtn.Name = "BeautifyTblBtn"
        Me.BeautifyTblBtn.ShowImage = True
        Me.BeautifyTblBtn.SuperTip = "1.连续表格美化;2.一字线填充空缺;3.删除末尾的句号"
        '
        'SplitTblBtn
        '
        Me.SplitTblBtn.Image = CType(resources.GetObject("SplitTblBtn.Image"), System.Drawing.Image)
        Me.SplitTblBtn.Label = "连续拆分"
        Me.SplitTblBtn.Name = "SplitTblBtn"
        Me.SplitTblBtn.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.SearchStdBtn)
        Me.Group2.Items.Add(Me.runBtn)
        Me.Group2.Items.Add(Me.AIwriting)
        Me.Group2.Label = "尊享功能"
        Me.Group2.Name = "Group2"
        '
        'SearchStdBtn
        '
        Me.SearchStdBtn.Image = CType(resources.GetObject("SearchStdBtn.Image"), System.Drawing.Image)
        Me.SearchStdBtn.Label = "标准查询"
        Me.SearchStdBtn.Name = "SearchStdBtn"
        Me.SearchStdBtn.ShowImage = True
        Me.SearchStdBtn.SuperTip = "查询国内外标准状态"
        '
        'runBtn
        '
        Me.runBtn.Image = CType(resources.GetObject("runBtn.Image"), System.Drawing.Image)
        Me.runBtn.Label = "一键懒人"
        Me.runBtn.Name = "runBtn"
        Me.runBtn.ShowImage = True
        Me.runBtn.SuperTip = "一键自动完成"
        '
        'AIwriting
        '
        Me.AIwriting.Image = CType(resources.GetObject("AIwriting.Image"), System.Drawing.Image)
        Me.AIwriting.Label = " AI 写作"
        Me.AIwriting.Name = "AIwriting"
        Me.AIwriting.ShowImage = True
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.DonateBtn)
        Me.Group4.Items.Add(Me.SettingBtn)
        Me.Group4.Items.Add(Me.AboutBtn)
        Me.Group4.Label = "关于"
        Me.Group4.Name = "Group4"
        '
        'DonateBtn
        '
        Me.DonateBtn.Image = CType(resources.GetObject("DonateBtn.Image"), System.Drawing.Image)
        Me.DonateBtn.Label = "捐赠"
        Me.DonateBtn.Name = "DonateBtn"
        Me.DonateBtn.ShowImage = True
        '
        'SettingBtn
        '
        Me.SettingBtn.Image = CType(resources.GetObject("SettingBtn.Image"), System.Drawing.Image)
        Me.SettingBtn.Label = "设置"
        Me.SettingBtn.Name = "SettingBtn"
        Me.SettingBtn.ShowImage = True
        '
        'AboutBtn
        '
        Me.AboutBtn.Image = CType(resources.GetObject("AboutBtn.Image"), System.Drawing.Image)
        Me.AboutBtn.Label = "关于"
        Me.AboutBtn.Name = "AboutBtn"
        Me.AboutBtn.ShowImage = True
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
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents StructureChkBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BibValidBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TermsChkBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BibRefChkBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BignumMdfBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents UnitMdfBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AboutBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BeautifyTblBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents SplitTblBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents MergeTblBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents CoverChkBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ForewordChkBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AbbChkBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents VarFontMdfBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ListChkBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SearchStdBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents runBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents DonateBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SettingBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ApplyStyleBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AIwriting As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
