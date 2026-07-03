# 2026年全国声学大会论文摘要集

基于 LaTeX 的学术会议论文集自动生成系统。从 Excel 摘要数据生成格式化的 PDF 论文集，包含超链接目录、作者单位自动编号、通讯作者标记等功能。

## 项目结构

```
.
├── main.tex                          # 单篇论文示例模板
├── main_full.tex                     # 完整论文集主模板
├── generate_tex.js                   # Excel 数据处理与 LaTeX 生成脚本
├── toc_content.tex                   # 生成的目录文件（自动生成）
├── papers_content.tex                # 生成的论文正文文件（自动生成）
├── 摘要信息.xls                       # 原始论文摘要数据
├── conference-proceedings-latex-skill/ # AI 辅助开发配置
└── .gitignore
```

## 依赖环境

- **Node.js** — 运行数据处理脚本，需要 `xlsx` 包
- **pdfLaTeX** — 需要支持 CJK（中文）编译，推荐 TeX Live 完整发行版
- **LaTeX 宏包**: `ctexart`、`newtxtext`、`newtxmath`、`hyperref`、`fancyhdr`、`geometry`、`longtable`、`setspace`、`indentfirst`

安装 Node.js 依赖：

```bash
npm install xlsx
```

## 构建流程

### 1. 生成 LaTeX 内容

从 Excel 数据生成论文正文和目录：

```bash
node generate_tex.js
```

此命令读取 `摘要信息.xls`，自动处理以下逻辑：

- 拆分并去重作者单位，按首次出现顺序编号
- 通讯作者姓名后添加星号标记
- 每位作者姓名后标注对应单位序号
- HTML 实体和特殊字符转义为 LaTeX 命令

### 2. 编译完整论文集

需要两次编译以更新交叉引用和目录页码：

```bash
pdflatex main_full
pdflatex main_full
```

输出文件：`main_full.pdf`

### 3. 单篇测试编译（可选）

用于验证模板样式：

```bash
pdflatex main
pdflatex main
```

输出文件：`main.pdf`

## 数据格式

Excel 文件（`摘要信息.xls`）应包含以下列：

| 列名 | 说明 |
|------|------|
| 论文专题 | 论文所属专题分类 |
| 论文编号 | 论文唯一编号 |
| 论文标题 | 论文题目 |
| 作者姓名1-12 | 作者姓名（最多12位） |
| 作者单位1-12 | 作者对应单位 |
| 通讯作者姓名 | 通讯作者姓名 |
| 通讯作者邮箱 | 通讯作者邮箱 |
| 摘要文本 | 论文摘要内容 |
| 论文关键字 | 论文关键词 |

## 排版特性

- 中文字号自定义命令（五号、小四、四号、三号、小二）
- 作者单位自动去重与编号
- 通讯作者星号上标标记
- 超链接目录，支持点状引导线
- 每篇论文独立页面，页脚显示通讯作者信息
- 页眉统一显示"2026年全国声学大会"

## 已知问题

目录中部分单行标题的引导线点可能无法正确延伸到右边界对齐页码，影响如 G0005 等编号的条目显示效果。

## 许可证

内部使用项目。
