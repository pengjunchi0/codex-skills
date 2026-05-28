# Codex 科研论文绘图 Skill

当前版本：`v1.0`

`Codex 科研论文绘图 Skill` 是一个面向科研论文配图工作流的 Codex Skill。它用于把参考图片、截图或生成图还原为 **Microsoft Visio `.vsdx` 原生可编辑图形**。核心目标不是把图片贴进 Visio，而是让 Codex 通过 Visio 原生形状、文本、连线、分组和样式重建论文配图、模型框架图、流程图和多面板科学图。


## 适用场景

适合：

- 根据 PNG/JPG/截图重建 Visio 图。
- 将 AI 生成的论文模型图转成可编辑 `.vsdx`。
- 按参考图修改已有 Visio 文件的布局、配色、字体或模块结构。
- 对复杂多面板科学图进行一比一结构复刻。
- 检查 `.vsdx` 是否误用了整张参考图嵌入。
- 给 Visio 图统一论文风格字体、配色和线条规范。

不适合：

- 只需要把图片插入 Visio 页面。
- 只需要普通图片编辑、抠图或美化。
- 不要求 Visio 原生可编辑性的纯位图复刻。

## 核心原则

最终交付的 `.vsdx` 应尽量由以下对象构成：

- Visio 原生矩形、圆形、线条、箭头、连接线。
- 可编辑文本。
- 可编辑分组。
- 原生近似绘制的小图表、热图、节点图、立方体、堆叠图。

禁止用整张参考图作为最终页面内容来冒充还原。参考图只能作为临时描摹依据；最终文件中不应留下完整的大尺寸参考 PNG/JPG。

## 仓库结构

```text
.
├── README.md
├── SKILL.md
├── agents/
│   └── openai.yaml
├── references/
│   └── rebuild-guidelines.md
└── scripts/
    ├── visio_page_tools.ps1
    └── visio_rebuild_scaffold.ps1
```

文件说明：

- `SKILL.md`：Codex Skill 的主入口，包含触发描述、工作流、验收标准和安全规则。
- `agents/openai.yaml`：Codex UI 元数据。
- `references/rebuild-guidelines.md`：复杂科学图还原准则，包括面板拆解、绘图顺序、风格参数和验证 rubric。
- `scripts/visio_page_tools.ps1`：辅助检查脚本，用于备份、导出预览、检查 `.vsdx` 包结构。
- `scripts/visio_rebuild_scaffold.ps1`：Visio 原生绘图脚手架，用于快速编写一比一重建脚本。

## 安装方式

### 作为 Codex Skill 使用

将本仓库克隆或复制到 Codex skills 目录。

Windows 示例：

```powershell
git clone https://github.com/pengjunchi0/codex-skills.git "$env:USERPROFILE\.codex\skills\visio-image-rebuilder"
```

或手动复制整个目录到：

```text
C:\Users\<你的用户名>\.codex\skills\visio-image-rebuilder
```

目录中必须保留：

```text
SKILL.md
agents/openai.yaml
references/rebuild-guidelines.md
scripts/*.ps1
```

安装后重启 Codex 或开启新会话，使 skill 被重新发现。

## 环境要求

推荐环境：

- Windows。
- Microsoft Visio。
- PowerShell。
- Git。
- Codex Desktop 或支持本地文件与工具调用的 Codex 环境。

说明：

- 完整 Visio 自动绘图依赖 Visio COM Automation，因此主要面向 Windows + Microsoft Visio。
- 不安装 Visio 时，仍可做 `.vsdx` 包结构检查或有限 XML 修改，但不适合完整一比一重建。

## 推荐使用方式

示例请求：

```text
使用 visio-image-rebuilder，根据这张参考图片重建 C:\path\model.vsdx，要求最终是 Visio 原生可编辑形状，不要整图嵌入。
```

```text
把这个 .vsdx 按参考图更换配色，保持布局不变，最终仍然可编辑。
```

```text
检查这个 .vsdx 是否只是嵌入了整张 PNG，如果是，请改成原生 Visio 形状重建。
```

## 标准工作流

### 1. 输入检查

Codex 应先确认：

- 目标 `.vsdx` 路径。
- 参考图片路径或已上传图片。
- 是否要求覆盖原文件。
- 是否需要保留原页面或清空重建。
- 是否允许打开 Visio。

然后执行：

- 备份原 `.vsdx`。
- 导出当前页面预览。
- 检查 `.vsdx` 包结构、页面数量、shape 数量和 `visio/media`。

### 2. 参考图拆解

对参考图做结构化拆解：

- 页面比例和边距。
- 主面板和子面板。
- 模块标题、编号、标题颜色。
- 主要数据流箭头。
- 虚线反馈、约束线、括号、图例。
- 重复元素，例如堆叠序列、卷积块、图节点、热图、表格、曲线图。
- 需要精确保留的文字和公式。

### 3. 原生形状重建

优先使用 Visio COM 自动化：

- `DrawRectangle`
- `DrawOval`
- `DrawLine`
- `Page.Export`
- `Shape.CellsU(...)`

使用统一坐标系统将参考图像素坐标映射到 Visio 页面英寸坐标。`scripts/visio_rebuild_scaffold.ps1` 已提供基础函数：

- `RectTL`
- `TextTL`
- `OvalTL`
- `LineTL`
- `DotTL`
- `VX`
- `VY`

### 4. 样式统一

建议默认论文图风格：

- 字体：`Times New Roman`。
- 主边框：`0.9-1.2 pt`。
- 内部边框：`0.5-0.8 pt`。
- 模块背景：低饱和浅色。
- 模块强调色：每个模块使用独立色系。
- 箭头：统一箭头样式。
- 虚线：用于反馈、约束和辅助连接。

### 5. 验证

完成后应检查：

- 是否有备份文件。
- 是否能导出预览。
- `.vsdx` 中是否存在大尺寸整图 PNG/JPG。
- 主要文本是否仍可编辑。
- shape 数量是否合理。
- 页面结构是否和参考图一致。

## 脚本说明

### `scripts/visio_page_tools.ps1`

用途：备份、导出预览、检查 `.vsdx` 包结构、关闭已打开的目标文档。

示例：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\visio_page_tools.ps1 `
  -VsdxPath "C:\path\model.vsdx" `
  -Backup `
  -InspectPackage
```

导出预览：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\visio_page_tools.ps1 `
  -VsdxPath "C:\path\model.vsdx" `
  -PreviewPath "C:\path\preview.png" `
  -ExportPreview
```

检查输出示例：

```text
Shape count: 845
Media entries: none
Large or raster media entries: 0
```

### `scripts/visio_rebuild_scaffold.ps1`

用途：创建完整重建脚本的起点。

建议做法：

1. 复制该脚本到工作区。
2. 修改 `Draw-ReferenceFigure`。
3. 使用参考图坐标逐步绘制面板、文本、箭头和重复元素。
4. 导出预览。
5. 检查 `.vsdx` 是否仍为原生形状。

示例：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\visio_rebuild_scaffold.ps1 `
  -VsdxPath "C:\path\model.vsdx" `
  -PageW 16 `
  -PageH 12 `
  -RefW 1448 `
  -RefH 1086 `
  -PreviewPath "C:\path\preview.png"
```

## 复杂科学图还原建议

对于复杂多面板模型图，不要直接从细节开始画。推荐顺序：

1. 页面尺寸和背景。
2. 大面板、分隔线、标题和编号。
3. 主流程箭头。
4. 文本框和模块框。
5. 重复视觉元素。
6. 小图表、热图、节点图、表格和公式。
7. 字体、线宽、配色统一。
8. 导出预览和包结构检查。


这个清单应作为绘图脚本的实现计划。

## 验收标准

一个合格的 Visio 还原结果应满足：

- 主体布局和参考图一致。
- 主要模块、标题、编号、箭头和说明文字齐全。
- 文字可编辑。
- 图形对象可单独选中和修改。
- 没有整张参考图作为最终底图。
- 配色、字体和线条风格统一。
- 有原文件备份。

## 常见失败模式

需要避免：

- 直接把整张 PNG 插入 Visio。
- 脚本超时后没有确认文件是否真的保存。
- 生成大量未分组形状，后续难以编辑。
- 全局替换颜色导致不同模块语义混乱。
- 字体变化后文字溢出。
- Visio 后台进程占用文件，导致保存失败。
- 未备份原文件。


当前目录结构：

```text
SKILL.md
agents/
references/
scripts/
README.md
```

如果未来要在同一仓库中维护多个 skill，建议调整成：

```text
skills/
  visio-image-rebuilder/
    SKILL.md
    agents/
    references/
    scripts/
```

## 状态

当前版本：第一版可用。

已包含：

- Skill 元数据。
- 复杂图重建流程。
- Visio COM 绘图脚手架。
- `.vsdx` 检查工具。
- 科学图一比一还原准则。

后续可以继续增强：

- 自动读取参考图尺寸。
- 自动生成面板坐标草图。
- 提供更多 motif helper，例如 cube、heatmap、graph、stacked sequence。
- 加入示例 `.vsdx` 和参考图测试用例。

## Version History

### v1.0 - Initial usable version

当前 v1 已完成：

- 建立 Codex Skill 基础结构：
  - `SKILL.md`
  - `agents/openai.yaml`
  - `references/rebuild-guidelines.md`
  - `scripts/visio_page_tools.ps1`
  - `scripts/visio_rebuild_scaffold.ps1`
- 明确核心规则：最终 `.vsdx` 应由 Visio 原生可编辑形状构成，不能用整张参考图嵌入冒充还原。
- 支持以参考图片为目标，对科研论文图进行结构化拆解。
- 支持复杂多面板论文图的绘图流程设计：
  - 总体框架区。
  - 子模块面板区。
  - 箭头、虚线反馈、约束连接。
  - 图节点、热图、堆叠序列、立方体等 motif 的原生近似绘制思路。
- 提供 Visio COM 绘图脚手架：
  - 页面尺寸设置。
  - 坐标转换。
  - 矩形、文本、圆形、线条、箭头、点的基础绘制函数。
  - 预览导出接口。
- 提供 `.vsdx` 检查工具：
  - 备份。
  - 导出预览。
  - 检查 shape 数量。
  - 检查 `visio/media` 中是否存在大尺寸整图图片。
- 提供科研图默认风格建议：
  - Times New Roman 字体。
  - 低饱和模块配色。
  - 主边框和内部边框线宽建议。
  - 模块语义配色。
- README 已加入中英文搜索关键词，便于 GitHub 和搜索引擎检索。

v1 的能力边界：

- v1 不是自动图像识别系统。它不会自动从图片中精准检测所有元素坐标。
- v1 主要提供 Codex 执行还原任务时的流程、约束、脚手架和验证工具。
- 完整一比一重建仍需要 Codex 根据参考图编写或调整绘图脚本。
- 完整 Visio 自动化依赖 Windows + Microsoft Visio。
- 对极小文字、复杂公式和高密度纹理，v1 推荐使用可编辑近似表达，并在交付时说明无法精确识别的部分。

### Planned v1.1

计划增强：

- 增加更多可复用 motif helper：
  - `DrawCube`
  - `DrawHeatmap`
  - `DrawGraph`
  - `DrawStackedSequence`
  - `DrawMiniChart`
- 增加参考图尺寸读取脚本。
- 增加 shape inventory 导出脚本，用于分析已有 `.vsdx` 的文本、颜色、位置和分组。
- 增加一个最小示例任务，展示从参考图到 `.vsdx` 的完整脚本结构。

### Planned v2.0

长期目标：

- 半自动参考图结构解析：
  - 面板检测。
  - 主色提取。
  - 文本区域估计。
  - 箭头和连接线候选检测。
- 生成初始 Visio 草图后，由 Codex 迭代修正。
- 支持更多输出格式：
  - SVG
  - PPTX
  - PDF
- 建立测试集，用多张科研图验证还原质量和可编辑性。
