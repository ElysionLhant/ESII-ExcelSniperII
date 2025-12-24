# Excel AI 智能助手 (VSTO 版)

这是一个基于 .NET Framework 4.7.2 开发的 Excel VSTO 加载项（Add-in），深度集成了 LLM（OpenAI、火山引擎等），旨在通过 AI 辅助自动化 Excel 数据处理任务。

##  核心功能

### 1. 智能填充与生成
- **选区捕获**：一键捕获当前 Excel 选区的截图和坐标信息。
- **多模态上下文**：
  - **文件拖拽**：支持拖拽文本文件（.txt, .csv, .json, .md）作为参考资料。
  - **图片支持**：直接拖拽图片文件（.png, .jpg 等），插件会自动将其作为视觉上下文发送给多模态大模型。
  - **手动输入**：提供文本框直接粘贴补充信息。
- **智能回填**：LLM 生成的结构化数据会自动填充回 Excel 表格，支持自动扩展行列。

### 2. Prompt (提示词) 管理
- **预设库**：内置常用任务 Prompt（如发票提取、数据清洗）。
- **自定义管理**：支持保存、加载和删除自定义 Prompt，配置持久化保存。

### 3. 宏 (Macro) 代码库 
- **VBA 管理**：在侧边栏直接管理常用的 VBA 宏代码。
- **一键运行**：无需打开 VBA 编辑器，直接点击运行即可执行选中的宏。
- **持久化存储**：宏代码库保存在本地，跨工作簿可用。
- *注意：运行宏功能需要在 Excel 信任中心开启信任对 VBA 工程对象模型的访问。*

### 4. 系统设置
- **API 配置**：支持自定义 OpenAI 兼容接口（Base URL, API Key, Model Name）。
- **持久化**：所有设置和预设均保存在本地 AppData 目录，重启 Excel 不丢失。

##  技术栈

- **框架**：.NET Framework 4.7.2, VSTO (Visual Studio Tools for Office)
- **语言**：C#
- **UI**：WinForms (使用 FlowLayoutPanel 实现自适应布局)
- **网络**：HttpClient (支持 TLS 1.2)
- **数据**：JSON 序列化存储

##  如何构建与运行

1. **环境要求**：
   - Windows 操作系统
   - Visual Studio 2019 或更高版本（需安装 "Office/SharePoint 开发" 工作负载）
   - Microsoft Excel

2. **打开项目**：
   - 使用 Visual Studio 打开 ExcelSP2/ExcelSP2.sln 解决方案文件。

3. **编译运行**：
   - 点击 Start 或按 F5。
   - Visual Studio 会自动编译并启动 Excel，加载项将出现在侧边栏或功能区中。

4. **宏功能配置**（如果使用宏库）：
   - 在 Excel 中，前往 **文件 > 选项 > 信任中心 > 信任中心设置 > 宏设置**。
   - 勾选 **信任对 VBA 工程对象模型的访问**。

##  使用指南

1. **配置 API**：首次使用请点击侧边栏底部的 " Settings"，填入你的 LLM API 信息。
2. **捕获数据**：选中 Excel 中的数据区域，点击 "Capture Selection"。
3. **添加素材**：将相关的参考文件或图片拖入 "Context Materials" 列表。
4. **选择指令**：在 Prompt 下拉框选择任务类型，或直接修改文本框内容。
5. **生成**：点击 "Generate & Fill"，等待 AI 处理并回填数据。
6. **运行宏**：在 "Macro Library" 区域选择或粘贴代码，点击 "Run" 执行自动化脚本。

##  常见问题

- **UI 显示不全？**：插件使用了自适应布局，尝试调整侧边栏宽度。
- **API 报错？**：请检查 API Key 是否过期，以及网络是否能连通 Base URL。
- **宏无法运行？**：请确保已开启 Excel 的 VBA 信任权限。
