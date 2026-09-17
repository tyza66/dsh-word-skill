# DSH Word Skill

一个将「阅读 Word 文档」与「精确修改 Word 内容和格式」两项能力统一封装的 [DSH](https://github.com/deepseek-ai/dsh) skill。

## 功能

- **读取** `.docx`：提取并返回结构化文本（摘要、问答、定向查找）。
- **精确编辑**：修改措辞、调整结构、修复样式/标题/表格/页眉页脚/间距等细节，并通过渲染验证确保结果无缺陷。
- **内置工具与示例**：附带通用版 Python 脚本，可直接用于渲染验证、文档审计和精确表格编辑。

## 结构

```text
dsh-word-skill/
|-- .dsh/
|   `-- skills/
|       `-- word/
|           |-- SKILL.md                入口：能力边界、工具契约、黄金路径
|           |-- README.md               skill 说明
|           |-- scripts/
|           |   |-- render_docx.py      通用版 DOCX 渲染器（LibreOffice headless，输出 page-N.png）
|           |   |-- word_audit.py       通用版文档审计器（标题层级、编号、直接格式覆盖、字体使用）
|           |   `-- table_geometry.py   通用版表格几何辅助（tblW/tblGrid/tcW 同步写精确列宽）
|           |-- examples/
|           |   |-- read_content.py     读取并转储段落、表格、页眉页脚、节几何
|           |   |-- edit_content.py     跨 run 和表格单元格的精确查找替换
|           |   |-- edit_tables.py      按内容权重设置列宽、单元格边距和可见边框
|           |   `-- tracked_changes.py  通过原始 OOXML 补丁添加修订标记
|           `-- references/
|               `-- format-precision.md 格式精度检查清单（间距、边框、对齐、表格布局）
```

## 安装

把 `.dsh/skills/word/` 放到你的项目根目录（含 `.git` 的目录）下的 `.dsh/skills/` 里即可；DSH 会自动发现 `SKILL.md`，无需重启。

## 工作原理

该 skill 是编排层，不重复实现底层能力，而是让 agent 通过 `bash` 工具直接运行 Python 完成读取与编辑：

- **读取**：用 `python-docx` 提取段落、表格、页眉页脚等结构化内容。
- **精确编辑**：用 `python-docx` 处理常规段落/样式/表格，用 OOXML 补丁处理批注、修订、超链接、域等高级能力。
- **渲染验证**：用 LibreOffice 渲染为 PNG 逐页检查。DSH 自身不内置文档阅读能力，因此查看 PNG 时优先用模型的视觉能力或图片读取工具；纯文本模型则借助视觉/OCR skill（如 `see`）；两者皆无时退回纯文本审计并明确告知用户。

## 使用方式

在 DSH 中用自然语言描述 Word 相关需求即可，例如：

- "帮我读一下这个 docx"
- "把第三章的标题改成黑体"
- "这个表格的列宽和边框调一下"

skill 会自动命中并按黄金路径执行：先读取掌握上下文，再编辑，然后渲染为 PNG 逐页检查，循环至无缺陷后交付最终 `.docx`。

脚本与示例可直接在 Python 环境中运行（先安装依赖）：

```bash
pip install python-docx lxml pdf2image

# 渲染验证
python3 .dsh/skills/word/scripts/render_docx.py input.docx --output_dir out/

# 编辑前审计
python3 .dsh/skills/word/scripts/word_audit.py input.docx

# 精确表格编辑
python3 .dsh/skills/word/examples/edit_tables.py input.docx output.docx
```

渲染需要本机安装 [LibreOffice](https://www.libreoffice.org/)（或通过 `SOFFICE_BIN` 指定路径）。

## 适用边界

- 适用于 `.docx` 的内容阅读与精确编辑。
- 不适用于创建电子表格、幻灯片、PDF 或手绘/插图类素材。

## 许可证

[MIT](./LICENSE)
