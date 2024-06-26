# StandardCheck AIMate

欢迎使用**StandardCheck AIMate**，一款专为标准文件形式检查设计的VSTO Add-in工具。无论你是科研工作者、技术文档编写者还是任何需要遵循标准文件格式的专业人士，**StandardCheck AIMate**旨在提升你的工作效率，确保你的文件严格遵循GB/T 1.1-2020等标准，展现出高度的专业性和严谨性。

![image](https://github.com/zhangzhangco/CheckStandardfile/assets/5515762/fbd1970b-f842-40f9-988b-dd57fccb9689)

## 特色功能

- **全面检查**：从文档结构到内容细节，全方位检查，确保符合标准。
- **智能提示**：对悬置段、冗余标题、缺失标题等问题提供智能修正建议。
- **自动纠错**：自动调整文献引用顺序、修正标点符号，验证引用有效性。
- **术语统一**：确保文件中的术语、缩略语使用一致，提升文档专业度。
- **格式规范**：量与单位的统一，变量斜体规范，以及数值的千位分隔等细节处理。
- **表格美化**：支持表格的连续合并，提供批量美化功能，提升文档的视觉效果。
- **一键操作**：“一键懒人”功能，简化操作流程，提升用户体验。

## 安装指南

1. 下载**StandardAI CheckMate**安装包。
2. 打开Microsoft Word，选择“文件” > “选项” > “加载项”。
3. 在“管理”下拉菜单中选择“COM加载项”，点击“转到”。
4. 点击“添加”，选择下载的安装包，完成安装。
5. 重新启动Microsoft Word，即可开始使用。

## 使用说明

1. 打开需检查的文档。
2. 点击Word菜单栏中的**StandardAI CheckMate**选项卡。
3. 选择相应的检查功能，如“封面检查”。
4. 根据提示修正文档中的问题。
5. 使用“一键懒人”功能，快速完成标准化处理。

![引用文件格式修正和标准查询](https://github.com/zhangzhangco/CheckStandardfile/assets/5515762/7e386dbd-ea04-416f-8bd8-c40cead79369)


![使用AI大模型处理段落](https://github.com/zhangzhangco/CheckStandardfile/assets/5515762/1e0740fd-7e60-4893-827b-31bf09acd8bb)

![使用AI大模型辅助编写](https://github.com/zhangzhangco/CheckStandardfile/assets/5515762/bff0ebf9-aa0d-4dbe-9a2e-3d2e9db78310)

## 完成情况
| 检查项目 | 检查内容 | 完成情况 |
| :---: | :--- | :---: |
| 封面 | 检查中英文文件名的大小写和空格 | √ |
| 正文文本 | 删除文档中的多余空行 | √ |
| 章节结构 | 检查悬置段、冗余标题和缺失标题 | √ |
| 一级标题存在性验证 | 确保每个一级标题都有相应的标题 | |
| 所有同级子条款标题一致性验证 | 确保同一层级的所有子条款标题一致，全有或全无 | |
| 标题末尾不应有句号验证 | 确保标题末尾不出现句号 | |
| 规范性引用文件的格式 | 检查排序和标点符号使用 | √ |
| 规范性引用文件 | 进行查询、验证和自动生成文本 | √ |
| 引用文件的提及 | 确保文中提及的 “规范性引用文件” 和 “参考文献” 应有相应章节 | √ |
| 术语格式 | 检查标点符号的使用 | √ |
| 术语的使用 | 确保术语在文中出现至少两次 | √ |
| 缩略语 | 缩略语应按字母顺序排序并在文中使用 | √ |
| 缩略语风格检查 | 检查缩略语使用是否规范，避免使用点号，考虑语言特定性 | |
| 列项数量验证 | 验证编号条款中是否有多于一个的列项 | |
| 列项深度验证 | 验证列项深度是否超过四层 | |
| 列项的标点符号 | 确保所有列项由适当的冒号或句号前置 | √ |
| 列项引导语 | 确保引导语存在且段落样式正确 | √ |
| 列项条目的结束标点 | 确保破碎句子的最后一项以句号结束，其他以分号或句号结束 | √ |
| 前言中的小节验证 | 前言中不应包含小节 | |
| 规范性引用中的小节验证 | 规范性引用中不应包含小节或引用 | |
| 节的存在性验证 | 文档必须包含 “范围”、“规范性引用” 和 “术语及定义” 部分 | √ |
| 节的顺序验证 | 验证文档节的顺序是否符合规范要求 | |
| 唯一子小节验证 | 验证是否有小节作为唯一子小节存在 | √ |
| 规范性参考文献项的风格验证 | 确保规范性参考文献项符合要求，特别是非标准出版物作为规范性引用的情况 | |
| 前言检查 | 前言中不应包含要求 | |
| 范围检查 | 范围中不应包含要求 | |
| 引言检查 | 引言中不应包含要求 | √ |
| 术语检查 | 术语中不应包含要求 | |
| 示例检查 | 示例中不应包含要求 | |
| 注释检查 | 注释中不应包含要求 | |
| 脚注检查 | 脚注中不应包含要求 | |
| 数字检查 | 4 位以上数字使用千位分隔 | √ |
| 百分比风格检查 | 百分比表示的公差应用括号括起来 | |
| 单位风格检查 | 数字与单位之间应有一定间隔，检查 SI 和非标准单位的使用 | √ |
| 数学变量 | 使用斜体，并采用 Times New Roman 字体 | √ |
| 文本美化 | 制作样式模板，确保中西文、数字与中文之间的间隔符合要求 | √ |
| 要素检查 | 对术语来源、公式、注释等要素进行检查 | |
| 表格 | 批量优化表格美观性，包括段落样式和去除多余的句号 | √ |
| 表格合并和拆分 | 批量合并和拆分表格 | √ |
| 一键执行 | 按顺序执行批量检查和处理 | √ |
| 样式重建 | 老SET转换为SET 2020 | |
	
## 常见问题

Q: 支持的Word版本有哪些？
A: 支持Microsoft Word 2016及以上版本。

Q: 如何更新至最新版本？
A: 访问GitHub项目页面下载最新版本的安装包，重复安装指南中的步骤即可更新。

## 贡献指南

欢迎各位专业人士、技术爱好者参与项目的贡献！如果你有好的想法或建议，请通过Issue提交或直接提交Pull Request。

## 许可证

本项目采用MIT许可证。详情请参见[LICENSE](LICENSE)文件。

## 联系方式

如有任何疑问或需要帮助，请通过项目的Issue功能提交问题，我们将尽快回复。

---
