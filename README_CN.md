# AI Excel 智能填充侧边栏

这是一个 Excel Web 加载项（Add-in），集成了 LLM（OpenAI、火山引擎等），可以根据上下文自动填充 Excel 表格。

## 功能特性

- **选区捕获**：捕获当前选中区域的坐标和截图。
- **上下文素材**：支持拖拽图片、文本文件（txt, csv, md, json）或 PDF。插件会读取文件内容并发送给 LLM 作为参考。
- **Prompt 管理**：保存、加载和管理自定义 Prompt 预设。内置了常用任务的默认预设。
- **真实 LLM 集成**：可配置 API 端点（兼容 OpenAI 协议）来生成数据。
- **自动扩展**：如果 LLM 生成的数据行数多于选区行数，插件会自动向下扩展 Excel 区域并保留格式。

## 如何运行

1. **前置条件**：
   - 已安装 Node.js（用于运行本地服务器）。
   - 拥有兼容 OpenAI 协议服务的 API Key（例如 OpenAI, 火山引擎/Ark）。

2. **启动本地服务器**：
   在当前文件夹打开终端并运行：
   ```bash
   npx browser-sync start --server --https --files "." --index "taskpane.html" --port 3000
   ```
   *注意：这将启动一个安全的本地服务器。你可能需要在浏览器中接受证书警告。*

3. **信任证书**：
   在浏览器中打开显示的 URL（例如 `https://localhost:3000/taskpane.html`）。你可能会看到安全警告。点击“高级” -> “继续前往 localhost（不安全）”以临时信任此会话。

4. **在 Excel 中旁加载 (Sideload)**：
   - 打开 Excel（桌面版或网页版）。
   - 转到 **插入 (Insert)** 选项卡。
   - 点击 **获取加载项 (Get Add-ins)**。
   - 点击 **我的加载项 (My Add-ins)**。
   - 选择 **上传我的加载项 (Upload My Add-in)**（通常在“管理我的加载项”下拉菜单下）。
   - 选择此文件夹中的 `manifest.xml` 文件。

5. **配置**：
   - 在侧边栏中，点击 **⚙️ Settings**。
   - 输入你的 **API Base URL**（例如 `https://api.openai.com/v1` 或你的火山引擎端点）。
   - 输入你的 **API Key**。
   - 输入 **Model Name**（例如 `gpt-4o` 或你的火山引擎端点 ID）。

6. **使用方法**：
   - 在 Excel 中选择一个区域。
   - 点击 **Capture Selection**。
   - 如果需要，拖拽参考文件（PDF/文本）。
   - **选择 Prompt**：从下拉菜单选择预设或输入自定义指令。你可以通过 "Manage" 按钮管理预设。
   - 点击 **Generate & Fill**。

## 自定义

- **图标**：`manifest.xml` 指向 `assets/` 中的占位符图标。你可以用真实的 PNG 图片替换它们。
