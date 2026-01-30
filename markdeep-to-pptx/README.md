# Markdeep Slides to PPTX Converter

将 Markdeep Slides 幻灯片转换为 PowerPoint (PPTX) 格式的工具，保留原始样式和布局。

## 简介

这个工具使用 Playwright 浏览器自动化来打开渲染后的 Markdeep Slides HTML 文件，提取幻灯片内容和样式，然后使用 PptxGenJS 生成高保真度的 PowerPoint 演示文稿。

## 特性

### 核心功能
- ✅ 自动解析 Markdeep Slides 渲染后的 HTML 结构
- ✅ 保留文本格式（粗体、斜体、下划线）
- ✅ 使用微软雅黑字体，支持中文显示
- ✅ 蓝色主题配色，与 Markdeep 默认主题一致

### 布局支持
- ✅ **顶部章节导航栏** - 显示所有章节，高亮当前章节
- ✅ **H2 标题下划线** - 蓝色装饰线
- ✅ **列表符号** - 蓝色圆点符号 (•)
- ✅ **页脚** - 章节标签和页码

### 元素支持
- ✅ 多级标题 (H1-H6)
- ✅ 有序/无序列表
- ✅ 表格
- ✅ Admonition 提示框 (note, tip, warning, error, question)
- ✅ 代码块
- ✅ 引用块
- ✅ 双栏布局

## 安装

```bash
cd markdeep-to-pptx
npm install
```

## 使用方法

### 基本用法

```bash
node src/index.js <input.html> [output.pptx]
```

### 示例

```bash
# 转换幻灯片，输出到默认位置 (output/<filename>.pptx)
node src/index.js ../markdeep-slides-project/Tutorial.html

# 指定输出路径
node src/index.js ../markdeep-slides-project/Tutorial.html ./my-presentation.pptx

# 快捷命令
npm run convert -- ../markdeep-slides-project/Example1.html
```

## 项目结构

```
markdeep-to-pptx/
├── package.json           # 项目配置和依赖
├── README.md              # 本文件
├── src/
│   ├── index.js           # 主入口 CLI
│   ├── slide-extractor.js # 使用 Playwright 提取幻灯片内容
│   ├── pptx-generator.js  # 生成 PPTX 文件
│   ├── debug.js           # 调试脚本
│   └── visual-debug.js    # 可视化调试（截图）
└── output/                # 默认输出目录
```

## 技术栈

- **Playwright** - 浏览器自动化，用于渲染和解析 HTML
- **PptxGenJS** - 生成 PowerPoint 文件
- **Node.js** - ES Modules

## 工作原理

1. **HTML 渲染**：使用 Playwright 无头浏览器打开 Markdeep Slides HTML 文件
2. **等待渲染**：等待 JavaScript 完全渲染幻灯片（包括 MathJax）
3. **DOM 提取**：从渲染后的页面中提取：
   - 幻灯片结构和类型（标题页、章节页、内容页）
   - 元素内容和格式（标题、列表、表格、Admonition）
   - 元素位置和尺寸
   - 导航栏章节信息
4. **PPTX 生成**：使用 PptxGenJS 创建对应的幻灯片：
   - 根据幻灯片类型选择不同渲染方式
   - 复刻导航栏和标题样式
   - 保持元素位置

## 支持的 Markdeep Slides 格式

本工具支持以下 Markdeep Slides 格式：

```markdown
**标题**
副标题

%%章节标题%%

# 小节标题

## 幻灯片标题

- 关键点1
- **粗体关键词**

> [!note] 提示标题
> 
> 提示内容
```

## 支持的元素类型

| 元素类型 | 支持状态 | 说明 |
|---------|---------|------|
| 标题页 | ✅ | `**标题**` 格式 |
| 章节过渡页 | ✅ | `# 小节标题` (H1) |
| 内容页标题 | ✅ | `## 幻灯片标题` (H2) + 下划线 |
| 段落 | ✅ | 保留格式 |
| 列表 (UL/OL) | ✅ | 蓝色圆点符号 |
| 表格 | ✅ | 保留表头样式 |
| Admonition | ✅ | note, tip, warning, error, question |
| 代码块 | ✅ | 等宽字体背景 |
| 引用块 | ✅ | 左边框样式 |
| 导航栏 | ✅ | 顶部章节导航 |
| 页脚 | ✅ | 章节标签 + 页码 |

## 样式特点

- **字体**：微软雅黑 (Microsoft YaHei)
- **主色调**：蓝色 (#2980B9)
- **标题**：蓝色加粗 + 下划线
- **列表符号**：蓝色圆点
- **Admonition**：彩色背景 + 左边框

## 限制

- 动画和过渡效果无法保留（PPTX 不支持 CSS 动画）
- 复杂的 SVG 图表需要单独处理
- MathJax 公式暂不支持
- 图片需要可访问的路径

## 开发

```bash
# 安装依赖
npm install

# 运行调试
node src/debug.js <input.html>

# 生成截图对比
node src/visual-debug.js <input.html>
```

## 许可证

MIT
