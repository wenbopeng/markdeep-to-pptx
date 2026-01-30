# Markdeep Slides to PPTX Converter

将 Markdeep Slides 幻灯片转换为 PowerPoint (PPTX) 格式的工具。

## 简介

这个工具使用 Playwright 浏览器自动化来打开渲染后的 Markdeep Slides HTML 文件，提取幻灯片内容和样式，然后使用 PptxGenJS 生成相应的 PowerPoint 演示文稿。

## 特性

- ✅ 自动解析 Markdeep Slides 渲染后的 HTML 结构
- ✅ 保留文本格式（粗体、斜体、下划线）
- ✅ 支持多级标题 (H1-H6)
- ✅ 支持有序/无序列表
- ✅ 支持表格
- ✅ 支持 Admonition 提示框 (note, tip, warning, error)
- ✅ 支持代码块
- ✅ 支持图片
- ✅ 支持双栏布局
- ✅ 保持原有的页面布局和位置

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
# 转换 Tutorial.html，输出到默认位置 (output/Tutorial.pptx)
node src/index.js ../markdeep-slides-project/Tutorial.html

# 指定输出路径
node src/index.js ../markdeep-slides-project/Tutorial.html ./my-presentation.pptx

# 快捷命令
npm run convert -- ../markdeep-slides-project/Example1.html
```

## 项目结构

```
markdeep-to-pptx/
├── package.json          # 项目配置和依赖
├── README.md             # 本文件
├── src/
│   ├── index.js          # 主入口
│   ├── slide-extractor.js # 使用 Playwright 提取幻灯片内容
│   └── pptx-generator.js  # 生成 PPTX 文件
└── output/               # 默认输出目录
```

## 技术栈

- **Playwright** - 浏览器自动化，用于渲染和解析 HTML
- **PptxGenJS** - 生成 PowerPoint 文件
- **Sharp** - 图像处理（如需要）

## 工作原理

1. **HTML 渲染**：使用 Playwright 打开 Markdeep Slides HTML 文件
2. **等待渲染**：等待 JavaScript 完全渲染幻灯片
3. **DOM 提取**：从渲染后的页面中提取结构化的幻灯片数据
4. **位置计算**：计算每个元素的准确位置和尺寸
5. **PPTX 生成**：使用 PptxGenJS 创建对应的幻灯片

## 支持的 Markdeep Slides 元素

| 元素类型 | 支持状态 | 说明 |
|---------|---------|------|
| 标题 (H1-H6) | ✅ | 保留大小和样式 |
| 段落 | ✅ | 保留格式 |
| 列表 (UL/OL) | ✅ | 支持嵌套 |
| 表格 | ✅ | 保留表头样式 |
| 图片 | ✅ | 支持本地和网络图片 |
| Admonition | ✅ | note, tip, warning, error |
| 代码块 | ✅ | 保留等宽字体 |
| 引用块 | ✅ | 带左边框样式 |
| 双栏布局 | ✅ | 保留列比例 |
| 高亮列表项 | ✅ | 支持五种颜色 |

## 限制

- 动画和过渡效果无法保留（PPTX 不支持 CSS 动画）
- 复杂的 SVG 图表可能需要单独处理
- MathJax 公式转换可能有精度损失

## 开发

```bash
# 安装依赖
npm install

# 运行测试转换
npm run test
```

## 许可证

MIT
