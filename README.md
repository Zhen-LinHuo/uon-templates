# University of Nottingham Presentation Templates

Brand-compliant presentation templates for the University of Nottingham — one for **PowerPoint / LibreOffice** and one for **LaTeX Beamer**, sharing the official colour palette and logo lock-up.

<div align="center">

[![License](https://badgen.net/github/license/Zhen-LinHuo/uon-presentation-templates)](https://github.com/Zhen-LinHuo/uon-presentation-templates/blob/main/LICENSE)
[![Stars](https://badgen.net/github/stars/Zhen-LinHuo/uon-presentation-templates)](https://github.com/Zhen-LinHuo/uon-presentation-templates)
[![Last commit](https://badgen.net/github/last-commit/Zhen-LinHuo/uon-presentation-templates)](https://github.com/Zhen-LinHuo/uon-presentation-templates/commits/main)

</div>

---

## Overview

Two ready-to-use templates that follow the University of Nottingham visual identity:

| Template | Directory | Output | Best for |
|---|---|---|---|
| **PowerPoint / LibreOffice** | [`powerpoint/`](powerpoint/) | `.pptx` / `.potx` / `.otp` | Talks where slides are edited in a GUI, shared with collaborators, or exported from PowerPoint / LibreOffice |
| **LaTeX Beamer** | [`beamer/`](beamer/) | `.pdf` (via XeLaTeX) | Talks written in markup, heavy maths, version-controlled decks, fully reproducible builds |

Both use the same brand palette — Nottingham Blue `#10263B` and Sky Blue `#009BC1` — with the official logo lock-up in the header band and cover, a progress bar, and a unified title grid.

**No system fonts are required.** The Beamer template draws every font from TeX Live itself (the `noto` package for Latin, Fandol for CJK), so a full TeX Live / Overleaf install builds it anywhere. The PowerPoint theme specifies Arial with Noto CJK fallbacks and embeds nothing fragile.

## Repository layout

```
.
├── powerpoint/                 # PowerPoint / LibreOffice template
│   ├── UoN-Presentation-Template.pptx   # sample deck (25 slides: layouts + charts + table)
│   ├── UoN-Presentation-Template.potx   # PowerPoint template
│   ├── UoN-Presentation-Template.otp    # LibreOffice template
│   ├── build.py                 # regenerates the deck from spec.json
│   ├── spec.json                # single source of truth: colours, fonts, geometry
│   └── assets/                  # logo lock-up and shield marks
├── beamer/                     # LaTeX Beamer template
│   ├── main.tex                 # the deck (input the title page, then write frames)
│   ├── titlepage.tex            # cover page — edit title / author / date here
│   ├── UoN.sty                  # theme: colours, header, progress bar, fonts
│   └── source/                  # logo images used by the theme
├── LICENSE
└── README.md
```

## PowerPoint / LibreOffice

### Use the template

- **PowerPoint**: open `UoN-Presentation-Template.potx` — PowerPoint creates a new untitled deck from it.
- **LibreOffice Impress**: open `UoN-Presentation-Template.otp` — Impress creates a new untitled presentation from it.
- **Just browse the sample**: open `UoN-Presentation-Template.pptx` to see every layout, 14 chart types, and a table.

The sample deck ships with these layouts and content:

- **Layouts** — Cover, Section Divider, Title and Content, Two Content, Comparison, Title Only, Blank.
- **Charts** — bar, column, line, area, pie, doughnut, scatter, radar, stacked column, percent-stacked column, waterfall, funnel, Pareto, histogram.
- **Special pages** — bold statement, quote, and closing ("Thank you").

### Rebuild from source

The deck is generated from `spec.json` (the single source of truth for colours, fonts and geometry) via `build.py`, which drives [OfficeCLI](https://github.com/) and then applies a few deterministic zip-layer fixes:

```bash
cd powerpoint
python3 build.py --step all       # build, validate, emit .potx, render preview
python3 build.py --step plan      # print the generated OfficeCLI commands only
python3 build.py --step potx      # regenerate the .potx from the .pptx
```

To rebrand, edit `spec.json` and rebuild — no scattered hard-coded colours.

## LaTeX Beamer

### Requirements

A **full TeX Live** (or Overleaf). Fonts resolve through TeX Live's own paths, so nothing is installed system-wide:

- **Latin** — the TeX Live `noto` package (Noto Sans / Serif / Mono).
- **CJK** — Fandol, via `ctex` with `fontset=fandol`.

Compile with **XeLaTeX**.

### Build

```bash
cd beamer
latexmk -xelatex main.tex
```

### Use the template

1. Edit **`titlepage.tex`** — set the title, subtitle, author, institute and date. This file is deliberately kept separate so reusers change only this one file.
2. Write your slides in **`main.tex`** — it loads `UoN.sty`, inputs the title page, then defines frames. Reuse the sample frames (columns, maths, code listing, table, TikZ figure) as blocks.
3. The theme provides the header band with logo, the progress bar, and section-divider slides automatically.

`cjktest.tex` is a minimal Chinese/English mixed-typesetting check — build it if you want to confirm CJK renders before writing a full deck.

## Customising the brand

Colours and fonts live in two places, both easy to edit:

- **PowerPoint** — `powerpoint/spec.json` (`colours`, `theme`). Rebuild with `build.py`.
- **Beamer** — `beamer/UoN.sty` (`\definecolor{NottinghamBlue}` / `{MalaysiaSkyBlue}` and the font block).

## Trademark notice

The University of Nottingham logos and marks in `powerpoint/assets/` and `beamer/source/` are the trademarks of the University of Nottingham and are **not** covered by this repository's Apache-2.0 licence. They are included so the templates render as designed. If you redistribute, confirm your use is permitted under the [University of Nottingham brand guidelines](https://www.nottingham.ac.uk/brand/).

## Credits

- **Beamer theme** — `beamer/UoN.sty` is derived from [Yanxi Jiang's University-of-Nottingham-Presentation-Template](https://github.com/YanxiJiang/University-of-Nottingham-Presentation-Template) (MIT), which in turn is based on a PolyU Beamer theme by Crumble Jon. Colour values follow the [official University of Nottingham visual identity](https://www.nottingham.ac.uk/brand/visual/colour.aspx). This repository's modifications are Apache-2.0; the inherited portions remain MIT.
- **Fonts** — [Noto](https://fonts.google.com/noto) (SIL Open Font License) and [Fandol](https://www.ctan.org/pkg/fandol) (CTAN), both distributed with TeX Live.
- **PowerPoint build** — generated with [OfficeCLI](https://github.com/) and `python-pptx`-style zip-layer fixes.

## Licence

Code and template sources are released under the [Apache License 2.0](LICENSE). The University of Nottingham logos remain the property of the University of Nottingham — see the trademark notice above.

---

# 诺丁汉大学演示模板

遵循诺丁汉大学视觉识别的品牌合规演示模板 —— 一套用于 **PowerPoint / LibreOffice**，一套用于 **LaTeX Beamer**，共享官方配色与校徽组合标识。

<div align="center">

[![License](https://badgen.net/github/license/Zhen-LinHuo/uon-presentation-templates)](https://github.com/Zhen-LinHuo/uon-presentation-templates/blob/main/LICENSE)
[![Stars](https://badgen.net/github/stars/Zhen-LinHuo/uon-presentation-templates)](https://github.com/Zhen-LinHuo/uon-presentation-templates)
[![Last commit](https://badgen.net/github/last-commit/Zhen-LinHuo/uon-presentation-templates)](https://github.com/Zhen-LinHuo/uon-presentation-templates/commits/main)

</div>

---

## 概述

两套开箱即用的模板，均遵循诺丁汉大学视觉识别：

| 模板 | 目录 | 输出 | 适用场景 |
|---|---|---|---|
| **PowerPoint / LibreOffice** | [`powerpoint/`](powerpoint/) | `.pptx` / `.potx` / `.otp` | 需在图形界面编辑、与协作者共享、或用 PowerPoint / LibreOffice 导出的演讲 |
| **LaTeX Beamer** | [`beamer/`](beamer/) | `.pdf`（经 XeLaTeX） | 用标记语言撰写、含大量公式、纳入版本控制、构建完全可复现的演讲 |

两者使用相同的品牌配色 —— 诺丁汉蓝 `#10263B` 与天蓝 `#009BC1` —— 页眉横条与封面带官方校徽组合标识、进度条以及统一的标题栅格。

**不依赖任何系统字体。** Beamer 模板的所有字体均取自 TeX Live 自身（拉丁字体用 `noto` 宏包，中日韩字体用 Fandol），因此任何完整 TeX Live / Overleaf 环境都能构建。PowerPoint 主题指定 Arial 并以 Noto CJK 兜底，不嵌入任何易失字体。

## 目录结构

```
.
├── powerpoint/                 # PowerPoint / LibreOffice 模板
│   ├── UoN-Presentation-Template.pptx   # 示例演示（25 页：版式 + 图表 + 表格）
│   ├── UoN-Presentation-Template.potx   # PowerPoint 模板
│   ├── UoN-Presentation-Template.otp    # LibreOffice 模板
│   ├── build.py                 # 依据 spec.json 重新生成演示
│   ├── spec.json                # 单一事实来源：配色、字体、几何
│   └── assets/                  # 校徽组合标识与盾徽
├── beamer/                     # LaTeX Beamer 模板
│   ├── main.tex                 # 演示正文（引入封面页后书写各帧）
│   ├── titlepage.tex            # 封面页 —— 在此修改标题 / 作者 / 日期
│   ├── UoN.sty                  # 主题：配色、页眉、进度条、字体
│   └── source/                  # 主题使用的校徽图片
├── LICENSE
└── README.md
```

## PowerPoint / LibreOffice

### 使用模板

- **PowerPoint**：打开 `UoN-Presentation-Template.potx` —— PowerPoint 会据此新建一份未命名演示。
- **LibreOffice Impress**：打开 `UoN-Presentation-Template.otp` —— Impress 会据此新建一份未命名演示。
- **仅浏览示例**：打开 `UoN-Presentation-Template.pptx`，可查看全部版式、14 种图表及一个表格。

示例演示包含以下版式与内容：

- **版式** —— 封面、章节分隔、标题和内容、双栏、对比、仅标题、空白。
- **图表** —— 条形图、柱状图、折线图、面积图、饼图、环形图、散点图、雷达图、堆积柱状图、百分比堆积柱状图、瀑布图、漏斗图、帕累托图、直方图。
- **特殊页** —— 粗体陈述、引用、以及致谢页（"Thank you"）。

### 从源码重建

该演示由 `spec.json`（配色、字体、几何的单一事实来源）经 `build.py` 生成；`build.py` 驱动 [OfficeCLI](https://github.com/)，随后施加若干确定性的 zip 层修补：

```bash
cd powerpoint
python3 build.py --step all       # 构建、校验、生成 .potx、渲染预览
python3 build.py --step plan      # 仅打印生成的 OfficeCLI 命令
python3 build.py --step potx      # 由 .pptx 重新生成 .potx
```

如需更换品牌，编辑 `spec.json` 后重建即可 —— 无需在各处硬编码配色。

## LaTeX Beamer

### 环境要求

**完整 TeX Live**（或 Overleaf）。字体经 TeX Live 自身路径解析，无需安装任何系统级字体：

- **拉丁字体** —— TeX Live 的 `noto` 宏包（Noto Sans / Serif / Mono）。
- **中日韩字体** —— Fandol，经 `ctex` 的 `fontset=fandol` 引入。

请使用 **XeLaTeX** 编译。

### 构建

```bash
cd beamer
latexmk -xelatex main.tex
```

### 使用模板

1. 编辑 **`titlepage.tex`** —— 设置标题、副标题、作者、单位与日期。该文件被刻意独立出来，使用者只需修改这一个文件。
2. 在 **`main.tex`** 中撰写各帧 —— 它载入 `UoN.sty`、引入封面页，然后定义各帧。可直接复用示例帧（分栏、公式、代码清单、表格、TikZ 图）作为积木。
3. 主题自动提供带校徽的页眉横条、进度条，以及章节分隔页。

`cjktest.tex` 是一个最小的中英混排检查文件 —— 在撰写完整演示前，可先构建它以确认中日韩字体渲染正常。

## 自定义品牌

配色与字体集中于两处，均易于修改：

- **PowerPoint** —— `powerpoint/spec.json`（`colours`、`theme`）。用 `build.py` 重建。
- **Beamer** —— `beamer/UoN.sty`（`\definecolor{NottinghamBlue}` / `{MalaysiaSkyBlue}` 及字体段）。

## 商标声明

`powerpoint/assets/` 与 `beamer/source/` 中的诺丁汉大学校徽与标识为诺丁汉大学的商标，**不受**本仓库 Apache-2.0 许可证约束。 它们被包含在内，仅为使模板按设计渲染。如需再分发，请确认您的用法符合[诺丁汉大学品牌指南](https://www.nottingham.ac.uk/brand/)。

## 致谢

- **Beamer 主题** —— `beamer/UoN.sty` 源自 [Yanxi Jiang 的 University-of-Nottingham-Presentation-Template](https://github.com/YanxiJiang/University-of-Nottingham-Presentation-Template)（MIT），后者又基于 Crumble Jon 的 PolyU Beamer 主题。配色数值遵循[诺丁汉大学官方视觉识别](https://www.nottingham.ac.uk/brand/visual/colour.aspx)。
- **字体** —— [Noto](https://fonts.google.com/noto)（SIL 开放字体许可证）与 [Fandol](https://www.ctan.org/pkg/fandol)（CTAN），均随 TeX Live 分发。
- **PowerPoint 构建** —— 由 [OfficeCLI](https://github.com/) 生成，并施加 `python-pptx` 风格的 zip 层修补。

## 许可证

代码与模板源码以 [Apache License 2.0](LICENSE) 发布。诺丁汉大学校徽仍归诺丁汉大学所有 —— 请参见上方商标声明。
