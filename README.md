# Markdeep Slides to PPTX Converter

将 Markdeep Slides 幻灯片转换为 PowerPoint (PPTX) 格式的工具，保留原始样式和布局。

## 简介

这个工具使用 Playwright 浏览器自动化来打开渲染后的 Markdeep Slides HTML 文件，提取幻灯片内容和样式，然后使用 PptxGenJS 生成高保真度的 PowerPoint 演示文稿。

## 安装

```bash
cd markdeep-to-pptx
npm install
npm link   # 注册全局命令 md2pptx
```

## 使用方法

### 基本用法

```bash
md2pptx <input.html> [output.pptx] [选项]
```

输出文件默认与输入文件**同目录同名**（扩展名改为 `.pptx`）。

### 选项

| 选项 | 说明 |
|------|------|
| `--clean` | 等价于同时开启 `--no-navbar --no-progressbar --no-chapter --no-page` |
| `--no-navbar` | 不渲染顶部章节导航栏 |
| `--no-progressbar` | 不渲染底部进度条 |
| `--no-chapter` | 不渲染左下角本章标签 |
| `--no-page` | 不渲染右下角页码 |

### 示例

```bash
# 基本转换（输出到 HTML 同目录）
md2pptx presentation.html

# 指定输出路径
md2pptx presentation.html ./output/my-slides.pptx

# 去掉导航栏和进度条
md2pptx presentation.html --no-navbar --no-progressbar

# 使用精简模式
md2pptx presentation.html --clean

# 去掉所有页脚元素
md2pptx presentation.html --no-progressbar --no-chapter --no-page
```

## 特性

### 布局与样式
- **主色调** `#034295` 蓝色，微软雅黑字体
- **顶部导航栏** — 显示所有章节，高亮当前章节（可关闭）
- **一级章节页** — 居左白字 + 贯通蓝色色块背景
- **H2 标题** — 蓝色加粗 + 全宽下划线
- **H3–H6 子标题** — 蓝色加粗，字号递减
- **底部进度条** — 按当前页进度填充（可关闭）
- **页脚** — 左下角本章标签 + 右下角页码（均可单独关闭）

### 元素支持

| 元素 | 说明 |
|------|------|
| 标题页 | `**标题**` 格式 |
| 章节过渡页 (H1) | 居左白字 + 蓝色贯通底纹 |
| 内容页 (H2) | 蓝色标题 + 下划线 |
| 结束页 / 致谢页 (`:::closing:::`) | 标题居中 + 说明文字，自动去掉导航栏、章节标签和页码，布局与标题页一致 |
| 段落 | 保留粗体 / 斜体 / 下划线，修正前置空行 |
| 有序 / 无序列表 | 蓝色符号，支持嵌套缩进 |
| 表格 | 三线表，按内容自动分配列宽 |
| Admonition 提示框 | note / tip / warning / error / question，低饱和配色，文字统一黑色 |
| 代码块 | 等宽字体 + 灰色背景 |
| 引用块 | 左边框样式 |
| 双栏布局 | 圆角背景、1.2 倍行距、文字位置修正 |

### Admonition 配色

| 类型 | 背景 | 边框 |
|------|------|------|
| note | `#d6eaf8` | `#7fb3d3` |
| tip | `#dcfad9` | `#a2f29a` |
| warning | `#ffe9d5` | `#ffc78f` |
| error | `#fde8e8` | `#f1948a` |
| question | `#f5eef8` | `#9b59b6` |

## 项目结构

```
markdeep-to-pptx/
├── package.json           # 项目配置和依赖（含 bin: md2pptx）
├── README.md
├── src/
│   ├── index.js           # CLI 入口，参数解析
│   ├── slide-extractor.js # Playwright 提取幻灯片内容
│   ├── pptx-generator.js  # PptxGenJS 生成 PPTX
│   ├── debug.js
│   ├── debug-navbar.js
│   ├── debug-html.js
│   └── visual-debug.js
└── output/                # 备用输出目录
```

## 技术栈

- **Playwright** — 浏览器自动化，渲染和解析 HTML
- **PptxGenJS** — 生成 PowerPoint 文件
- **Node.js** — ES Modules

## 工作原理

1. **HTML 渲染**：Playwright 无头浏览器打开 Markdeep Slides HTML
2. **等待渲染**：等待 JavaScript 完全渲染（含 MathJax）
3. **DOM 提取**：提取幻灯片结构、元素内容与格式、位置尺寸、导航信息
4. **PPTX 生成**：根据幻灯片类型选择渲染方式，复刻样式与布局

## 限制

- 动画和过渡效果无法保留
- 复杂 SVG 图表需单独处理
- MathJax 公式暂不支持
- 图片需要可访问的本地路径

## 许可证

MIT
