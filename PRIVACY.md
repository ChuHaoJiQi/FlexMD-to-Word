# FlexMD-to-Word Privacy Policy / 隐私政策

This plugin converts Markdown text into Word (.docx) files with style controls. We respect your data privacy and describe below how data is handled.

## EN

### Data Processed
- Input: Markdown string and optional style parameters (e.g., fonts, sizes, style_profile, filename).
- Output: Generated Word `.docx` file returned to Dify.

### Storage & Retention
- The plugin does not persist user content. Generated files are streamed back to Dify and not stored by the plugin beyond the execution lifecycle.
- Logs: The plugin itself does not log input content. Any logs are limited to runtime errors (e.g., missing dependencies).

### External Transfers
- No external credentials are required. The plugin does not send user content to third-party services.
- Note: If your Markdown references external images (e.g., `![](https://...)`), underlying libraries may fetch these images to embed them in the `.docx`. This is a direct retrieval, not a data exfiltration.

### Security
- The plugin validates runtime dependencies (`markdown`, `html2docx`).
- Filenames are sanitized to ensure `.docx` extension; no arbitrary file writes outside the plugin’s sandbox.

### Contact
- Author: `chuhaojiqi` (GitHub: https://github.com/chuhaojiqi)

### Changes to Policy
- This policy may be updated alongside plugin releases. Major changes will be noted in `CHANGELOG.md`.

## 中文（简体）

### 处理的数据
- 输入：Markdown 文本与可选样式参数（如字体、字号、样式预设、输出文件名）。
- 输出：生成的 Word `.docx` 文件并返回至 Dify。

### 存储与保留
- 插件不持久化用户内容。生成文件仅在执行过程中存在，并在返回结果后不保留。
- 日志：插件不记录输入内容，仅输出运行时错误（如依赖缺失）。

### 外部传输
- 无需任何外部凭据。插件不会将用户内容发送至第三方服务。
- 说明：若 Markdown 中引用远程图片（如 `![](https://...)`），底层库可能会为嵌入 `.docx` 而下载该图片。这属于直接资源获取，而非数据外传。

### 安全
- 插件会校验运行时依赖（`markdown`、`html2docx`）。
- 输出文件名会被规范化为 `.docx`；不会在插件沙箱之外进行任意文件写入。

### 联系方式
- 作者：`chuhaojiqi`（GitHub：https://github.com/chuhaojiqi）

### 政策变更
- 本政策可能随插件版本更新而调整。重大变更将记录于 `CHANGELOG.md`。