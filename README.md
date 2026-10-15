# business-report-ppt · 商务汇报 PPT 模板化装配

用**代码装配**而不是手拖：内容 = outline 声明，模板 = 用户自备的现货 `.pptx`，引擎 = donor 深拷贝 + 槽位替换，每页可预览校验。

本技能的知识（风格族 / 配色 / 字体阶梯 / 页面解剖 / 构图词表）全部从 **91 套欧美扁平风商务模板实测蒸馏**；模板文件不在本仓库内（BYO，用户自备），引擎对任何 `.pptx` 通用。

## 特性

- **模板化装配引擎**：`analyze` 提取文本槽位 → `build` donor 深拷贝 + 槽位替换（保留原格式）+ `drop` 删模板 logo + `fit` 长文本缩字号 → `verify` 查占位残留，全链路 CLI 化
- **11 个风格族设计规范**：商务红黑 / 蓝色商务 / 黑金高端 / 深色杂志 / 灰白简约 / 清新文艺 / 蓝绿科技 / 互联网科技 / 橙黄商务 / 黄色明快 / 红色商务，各带实测主色表
- **字体与版式阶梯**：微软雅黑主导；封面主标 66–115pt / 页标题 28–36 / 正文 12–16（投影下限 12pt）
- **页面解剖**：封面 → 目录 → [过渡页 + 内容×3–6] × N章 → 结尾，过渡页是节奏灵魂
- **自带原创合成模板**：`make_demo_template.py` 从零生成（商务红黑 / 蓝色商务 / 清新文艺三主题），没有模板库也能跑通全链路
- **PowerPoint 预览校验**：`export_preview.py` 逐页导出 PNG 目检（需本机 PowerPoint）
- **可选付费闸口**：`payment_gate.py` 实现支付宝 AI 按量付费（A2M）契约本地侧（402 账单 / Proof 验签 / 24h 幂等 / 履约确认）

## 快速开始

```bash
# 依赖
pip install python-pptx          # 引擎核心
pip install pywin32              # 仅预览导出需要(另需本机 Microsoft PowerPoint)

# 1. 无模板库?先生成一个原创合成模板练手(三主题可选)
python scripts/make_demo_template.py assets/demo_template.pptx --style redblack

# 2. 分析模板:认页型、提取文本槽位
python scripts/build_deck.py analyze "<你的模板.pptx>" --out slots.json

# 3. 参照 samples/demo_outline.json 写 outline(donor + slots + drop + fit)

# 4. 装配 + 校验
python scripts/build_deck.py build "<模板>" outline.json out.pptx
python scripts/build_deck.py verify out.pptx          # 占位残留必须清零

# 5. 预览目检(有 PowerPoint 时)
python scripts/export_preview.py out.pptx preview/
```

## 模板库自备（BYO）

本仓库**不含第三方模板文件**。三条路，按场景选：

1. **有模板库**：`export BUSINESS_PPT_ROOT="/path/to/模板目录"`，然后 `python scripts/template_lib.py pick 商务红黑` 检索同源欧美扁平风模板包
2. **只有零散模板**：跳过检索，直接对候选 `.pptx` 跑 `analyze` 认页——引擎不依赖索引
3. **什么都没有**：用仓库自带的 `assets/demo_template.pptx` 先跑通全链路

## 装配纪律（速览）

- 字数对齐：新文本 ≈ 原文本（汉字×2、拉丁×1 折算）；≥60pt 槽位中文 ≤4 字
- 品牌清理：模板自带 logo、版权页、论坛地址一律 drop 或弃用该页
- 只改 outline 重跑，不手改 pptx；业绩数据红=涨、绿=跌（中国惯例）
- 完整规范见 [SKILL.md](SKILL.md) 与 `references/`

## 目录结构

```
business-report-ppt/
├── SKILL.md                        # 技能主文件:五步工作流 + 纪律 + 付费契约
├── references/
│   ├── 设计系统.md                  # 网格/字体阶梯/11 风格族配色/页面解剖/构图词表
│   ├── 页面版式模式.md              # donor 认页/outline 写法/替换纪律/校验清单/FAQ
│   └── 模板索引.md                  # 91 套同源模板索引 + 风格族推荐表 + 格式警示
├── scripts/
│   ├── templates.json              # 模板库索引元数据(编号/文件/风格族/明暗/主色板/格式)
│   ├── template_lib.py             # 模板库检索:families/list/pick/path/check(BYO 解析)
│   ├── build_deck.py               # 装配引擎:analyze/build/verify
│   ├── export_preview.py           # PowerPoint COM 每页导出 PNG 目检
│   ├── make_demo_template.py       # 原创合成模板生成器(三主题)
│   └── payment_gate.py             # 付费闸口(A2M 契约本地侧,可选)
├── assets/
│   └── demo_template.pptx          # 原创合成模板:封面/目录/过渡/卡片/数据/结尾
└── samples/
    ├── demo_outline.json           # 完整示例:8 页《季度业务汇报》(donor 复用演示)
    └── demo_季度汇报.pptx           # 按上述 outline 装配的成品(已预览校验)
```

## 效果案例

**① 年终述职（商务红黑 · 完整工作流）**
> **用户**：下周要做年终述职，帮我做一份年终总结 PPT，正式稳重一点的。
> **Agent**：好，开工前确认三件事：① 汇报场合和听众？② 公司有指定模板吗？③ 有没有现成的业绩数据？
> **用户**：向管理层述职，没有指定模板。全年营收 2.4 亿、同比增长 22%、毛利率 33.5%……
> **Agent**：场合=正式汇报 → 定**商务红黑**风格族。先立内容框架：封面 / 目录 / 三章 / 结尾……确认后开始装配。
> **Agent**：已执行 analyze 认页 → outline → build → verify 残留清零 → 逐页预览。模板自带锐普 logo 已 drop、版权论坛页已弃用。交付《2026年度工作总结汇报》（10 页）。

**② 项目评审答辩（蓝色商务 · 用户自备模板库）**
> **用户**：明天上午项目评审答辩，我 D 盘有个蓝色系的商务模板文件夹，用那套帮我做份智慧园区项目方案 PPT。
> **Agent**：识别到 12 个模板，其中 9 个是旧版 .ppt 二进制格式（不能直接构建），已自动跳过，选定 OOXML 候选。蓝色系+答辩场合 → 推荐 `027蓝绿科技` 做 donor。
> **用户**：核心是建设背景、总体架构、实施计划三块，投资概算 3800 万，工期 14 个月。
> **Agent**：框架：背景（2 页）→ 架构（3 页）→ 计划（2 页）+ 投资概算数据页。build/verify/预览全过，8 页成品交付；里程碑拆季度？只改 outline 重跑即可。

**③ 新品上线宣传（清新文艺 · 无模板库快速兜底）**
> **用户**：新产品下周上线，做份内部宣传 PPT，风格年轻清新点，今天就要。
> **Agent**：急活走快速链路。没有自备模板库 → 先用技能内置的原创合成模板保交付；之后想要正式模板，同一份 outline 换 donor 重跑，不用重写。
> **用户**：产品是面向 Z 世代的便携咖啡机，卖点：3 分钟手冲级口感、胶囊便携、马卡龙配色，定价 399。
> **Agent**：6 页框架已交付。slogan 想换？改一行 outline，10 秒重跑出图。

## 许可

[MIT](LICENSE) © 2026 设备人
