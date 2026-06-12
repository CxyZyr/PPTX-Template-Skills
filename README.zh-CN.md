# PPTX 模板技能

[English](README.md) | **简体中文**

两个可组合的技能，把一份 PowerPoint **模板**转化为一份全新填充的演示文稿——
全程无需智能体手动从幻灯片上读取形状索引。解析阶段产出一份可移植、机器可读的契约；
生成阶段按页填充该契约，同时保留模板原有的视觉逻辑。

```
template.pptx ─[pptx-template-parsing]→ spec.json ─[ppt-template-adaptation]→ filled deck.pptx
```

## 两个技能

| 技能 | 职责 | 产出 |
|---|---|---|
| [`pptx-template-parsing`](skills/pptx-template-parsing) | **解析**：把模板解析成语义契约 | `spec.json` + `summary.txt` + 可选的 `renders/*.png` |
| [`ppt-template-adaptation`](skills/ppt-template-adaptation) | **生成**：依据 `spec.json` + `content_plan.json` 生成新文稿 | 填充后的 `deck.pptx` + 渲染图 |

解析器会对每个形状分类（标题 / 正文 / 卡片 / 图标 / logo / 图片 / 图表 /
表格），推断每页的角色，检测重复的卡片结构，并输出逐页的 `fill_plan` 槽位映射。
生成器读取该槽位映射，先填文本，再填图标、logo 与图片——
始终保持几何位置、对齐方式、主题色以及段落/文本块结构不变。

## 为什么先解析

解析器的输出即契约。生成阶段绝不凭记忆或上一份失败稿去挑选 PowerPoint 形状索引。
若某个槽位、卡片数量或样式有误，修复应落在解析侧——而不是在生成侧堆叠针对个例的兜底逻辑。
这样能让两个技能保持可移植，也让故障模式可被诊断。

## 快速开始

环境要求：Python 3.12+、[`python-pptx`](https://python-pptx.readthedocs.io/)。渲染需要
PATH 上有 `libreoffice` + `pdftoppm`（尽力而为；缺失时解析器仍会产出 `spec.json`）。
语义图标通过 `rsvg-convert` 使用 Bootstrap Icons。可选的网络配图需要 Tavily API key
（`TAVILY_API_KEY`）；没有时则保留模板原图。

```bash
# 1. 把模板解析成 spec.json
python3 skills/pptx-template-parsing/scripts/parse_template.py \
  --template "path/to/template.pptx" \
  --out workspace/specs/my_template

# 2. 依据 spec 生成内容计划骨架
python3 skills/ppt-template-adaptation/scripts/scaffold_content_plan.py \
  --spec workspace/specs/my_template/spec.json \
  --out workspace/plans/my_deck.content_plan.json \
  --output-pptx workspace/out/my_deck.pptx

# 3. 编辑计划：填写 keep_slides、outline 与 pages[]

# 4. 校验全局文本一致性
python3 skills/ppt-template-adaptation/scripts/validate_content_plan.py \
  --plan workspace/plans/my_deck.content_plan.json

# 5. 先逐页填充（增量），再做整体构建
python3 skills/ppt-template-adaptation/scripts/apply_content_plan.py \
  --plan workspace/plans/my_deck.content_plan.json --pages 0,1 --render

python3 skills/ppt-template-adaptation/scripts/apply_content_plan.py \
  --plan workspace/plans/my_deck.content_plan.json --render
```

## 仓库结构

```
skills/
  pptx-template-parsing/      # template.pptx -> spec.json
    SKILL.md                  # 何时用、如何用
    scripts/                  # parse_template、shape_classifier、slide_classifier
    references/               # spec 模式、幻灯片角色、分类信号
    lib/pptx_toolkit/         # 形状遍历、几何、渲染辅助
  ppt-template-adaptation/    # spec.json + content_plan.json -> 填充后的文稿
    SKILL.md
    scripts/                  # scaffold / validate / apply_content_plan、render、tavily
    references/               # workflow、text-coherence、icon-fill、image-fill
    lib/pptx_toolkit/
```

每个技能都刻意做成自包含：各自携带一份 `lib/pptx_toolkit`，以便独立使用。
完整的工作流与参考阅读顺序见各技能的 `SKILL.md`。

## 许可证

[MIT](LICENSE) © 2026 JerryChou
