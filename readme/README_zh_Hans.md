## FlexMD-to-Word

### 描述（中文）
- 将 Markdown 内容一键转换为高质量的 Word（.docx）。
- 支持标题层级、列表、表格、图片、代码块以及段前/段后间距等样式控制。
- 无需外部凭据，仅依赖运行时库（markdown、html2docx）。
- 适用于 Dify 工作流，Python Runner 3.12。

### 使用方法
- 必填：`markdown`（字符串，待转换的 Markdown 文本）。
- 可选：`filename`（默认 `document.docx`）、`style_profile`（如：学术论文/商务报告/技术文档），以及字体/字号覆盖项。
- 输出：Word `.docx` 文件。

### 其他
- 图标：`_assets/icon.svg` / `_assets/icon-dark.svg`。
- 不需要 OAuth 或任何外部凭据。
