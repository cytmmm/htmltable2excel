## HTML Table to Excel（Chrome 扩展，MV3）

[English](README.md)

一键将网页中的 HTML 表格导出为真实的 Excel（.xlsx）。在表格附近右键即可导出；适配常见"表头/表体分离"的双表结构，会自动合并后导出；所有处理均在本地完成。

### 功能特点
- 右键菜单：在右键处就近选择最近的 `<table>` 导出
- 自动合并：兼容表头和数据拆分为两张表的场景
- 真·XLSX：在浏览器中生成 OOXML 并打包 ZIP（无后端、无外部依赖）
- 隐私友好：不收集、不上传任何数据
- **超链接支持**：单个链接转换为 Excel HYPERLINK 公式，多个链接追加为文本
- **图片支持**：提取 alt/title 和 src 属性，将图片信息追加到单元格文本

### 开发者模式安装
1. 克隆或下载本仓库
2. 打开 Chrome → `chrome://extensions/` → 打开"开发者模式"
3. 点击"加载已解压的扩展程序"，选择项目文件夹

### 使用方法
- 打开包含 HTML `<table>` 的网页
- 在目标表格附近点击右键 → "导出此处的表格为 Excel"
- 表格会短暂高亮，浏览器下载 `.xlsx` 文件

### 权限说明
- `contextMenus`：注册右键菜单
- `activeTab` + `scripting`：仅在用户触发导出时向当前页注入脚本
- 为便利声明了 content script，但不做后台抓取

### 隐私
- 详见 `PRIVACY.md`。不收集、不存储、不共享任何用户数据；所有处理均在本地内存完成。

### 开发
- 后台 Service Worker：`background.js`
- 内容脚本：`content.js`
- Manifest：`manifest.json`
- 图标工具：打开 `tools/icon_generator.html` 生成多尺寸 PNG，放到 `icons/`

热更新提示：
- 每次修改后在 `chrome://extensions/` 刷新扩展，然后刷新目标页面
- 内容脚本日志在 DevTools Console 中，前缀为 `[Table2Excel]`

### 许可协议
MIT — 见 `LICENSE`。
