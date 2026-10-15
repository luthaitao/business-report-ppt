---
name: business-report-ppt
display_name: 商务汇报PPT
display_name_en: Business Report PPT
description: 用 python-pptx 程序化装配商务汇报/年终总结/产品介绍/企业宣传类 PPT:模板库(用户自备)选型 + donor 深拷贝槽位替换引擎 + PowerPoint 预览校验。当用户要求做商务汇报 PPT、用模板生成 PPT、套用某风格(商务红黑/蓝色商务/黑金/简约/清新)出幻灯片、或对既有汇报 PPT 改版重排时使用。
description_zh: 商务汇报PPT模板化装配:自备模板库选型 + donor 槽位替换引擎 + 预览校验,设计规范从 91 套欧美扁平风模板实测蒸馏
description_en: "Template-driven business report PPT assembly: BYO template library, donor-copy slot replacement engine, PowerPoint preview verification; design rules distilled from 91 flat-style business templates."
version: 1.0.0
author: 设备人
whenToUse: 当任务涉及商务汇报/年终总结/产品展示/企业宣传类 PPT 的生成或改版,需要套用商务模板风格、按模板版式装配幻灯片、或从模板库选型时
agent_created: true
---

# 商务汇报 PPT 模板化装配

用代码装配而不是手拖:**内容 = outline 声明,模板 = 用户自备的现货 .pptx,引擎 = donor 深拷贝 + 槽位替换,每页可预览校验**。
本技能的知识(风格族/配色/字体/版式)全部从模板库实测蒸馏;模板文件不在技能内(BYO,见下),引擎对任何 .pptx 通用。

## 何时使用 / 不使用

- ✅ 商务汇报、年终/季度总结、产品展示、企业宣传、竞聘报告、毕业答辩
- ✅ 用户指定或暗示了风格("正式一点""黑金高端""简约清新""红色汇报")
- ✅ 对既有汇报 PPT 换模板改版重排
- ❌ 公司官方合规模板(如企业 VI 模板)→ 用对应的公司模板技能,不走本引擎
- ❌ 只要 5 分钟急活初稿 → 直接让用户去 WPS/Kimi 套模板更快
- ❌ 需要复杂图表(SmartArt 级数据图)→ 本引擎做文本槽位替换,图表页优先选模板自带图表页做 donor

## 模板库自备(BYO)——先读这条

本技能不含模板文件。三条路,按场景选:

1. **有模板库**:把目录设为环境变量 `BUSINESS_PPT_ROOT`(推荐,换机器不用改文件),或填 `scripts/templates.json` 的 `root` 字段。若自备的是同源的欧美扁平风模板包,`template_lib.py` 的风格族检索可直接用。
2. **只有零散模板**:跳过模板库检索,直接对候选 `.pptx` 跑 `analyze` 认页——引擎不依赖索引。
3. **什么都没有**:用技能自带的合成模板 `assets/demo_template.pptx`(原创扁平风,六类页型齐全)先跑通全链路,或运行 `python scripts/make_demo_template.py` 重新生成一个改着用。

## 环境依赖(开工前检查)

1. Python 3.x + `python-pptx`(`pip install python-pptx`);预览导出另需 `pywin32` + 本机安装 Microsoft PowerPoint
2. 无 PowerPoint 时跳过预览,靠 `verify` + 结构自检
3. 配了模板库才需要:`python scripts/template_lib.py check` 看状态(未配置会打印 BYO 用法,不是错误)

## 付费调用契约(Pay Skill · 支付宝 AI 按量付费 A2M)

本技能为**付费技能(Pay Skill)**,按次计费。单次调用价格与 SkillHub 发布表单定价一致(当前 0.1 元/次;调整定价必须同步修改本说明与 `scripts/payment_gate.py` 的 `PRICE_CNY`)。调用遵循支付宝 AI 按量付费契约,由平台调度层与本技能脚本协同完成 **probe → pay → complete → ack** 四步:

| 阶段 | 责任方 | 行为 |
|---|---|---|
| probe(探测) | 平台调度层 | 调用前探测用户是否已持有本技能的有效付费凭证(`python scripts/payment_gate.py probe`);已付费直接放行 |
| 402 Payment-Needed(账单下发) | 技能脚本 | 未付费时向标准输出返回 HTTP 402 与 `Payment-Needed` 账单头(金额 = 发布定价),暂停交付,平台据此唤起支付宝收银 |
| pay(支付) | 支付宝 | 用户本人扫码/免密完成付款;Agent 只编排,不代付,不接触任何密钥 |
| complete(扣费确认) | 平台调度层 | 扣费成功后向重试请求下发 `Payment-Proof`(付费凭证) |
| ack(验签放行) | 技能脚本 | 携带 Payment-Proof 重试时执行 `python scripts/payment_gate.py verify --proof '<凭证>'` 验签;验付复用支付宝官方能力(`alipay.aipay.agent.payment.verify`),无效/过期凭证拒绝服务,结果不明不得要求重复付款 |

合规五要点(平台审核逐项检查):

1. **402 账单下发**:账单金额必须与发布定价一致,不得误导用户
2. **携带凭证重试**:用户付款后,`Payment-Proof` 随重试请求下发到技能侧
3. **验调用**:verify 验签通过才放行主流程;无效/过期凭证一律拒绝服务
4. **履约确认**:交付完成后执行 confirm(ack);失败可重试且保持幂等
5. **订单持久化与幂等**:订单与付费状态缓存在 `.pay_cache/state.json`;同一 `Payment-Proof` 在 24h 幂等窗口内仅生效一次——不重复扣费、不重复交付

免费场景(本地直接跑脚本、未经过平台计费调度)闸口自动放行,不影响调试与自测。

## 五步工作流(严格遵守)

### 第 1 步:内容框架先行(每页双栏)

先把每页列成清单,**每页两栏**:
- **内容栏**:页型(封面/目录/过渡/内容/数据/结尾)+ 标题(结论句)+ 要点 + 数字(注明口径,不确定写"待补",禁止编造)
- **构图栏**:内容语义 → 构图声明(查 `references/设计系统.md` 第 5 节构图词表:并列/层级/递进/扩散/聚合/包含/组织/联动/循环/流程/数据/图文)
- 节奏:封面 → 目录 → [过渡 + 内容×3–6] × N章 → 结尾;相邻页构图不得重复;每章至少 1 页数据页

### 第 2 步:选风格族 + 选模板

1. 按场合查 `references/模板索引.md` 的风格族推荐表(正式汇报→商务红黑;机关答辩→蓝色商务;高端→黑金;通用→灰白简约;文艺→清新文艺;科技→蓝绿/互联网;产品→橙黄)
2. 自备同源模板包:`python scripts/template_lib.py pick <风格族>` 拿首选;只有零散模板:直接对候选文件跑 analyze
3. **格式红线**:只选 OOXML(`.pptx`,含个别 PK 头的假 .ppt);旧二进制 .ppt 需先用 PowerPoint/WPS 另存为 .pptx,否则换一个能直接构建的
4. 4:3 模板只用于 4:3 场合

### 第 3 步:analyze 认页

```bash
python scripts/build_deck.py analyze "<模板路径>" --out slots.json
```
按 `type` 认页(封面/目录/过渡/内容/数据/结尾),记下每类页型的 donor 页码与槽位 id;**版权/论坛宣传页弃用**。认不准就对该模板跑 `export_preview.py` 看图。

### 第 4 步:写 outline 并装配

付费调用先过闸口:`python scripts/payment_gate.py gate`(未付费返回 402 账单,按上文契约处理;本地/免费调试自动放行)。

参照 `samples/demo_outline.json`(配 `assets/demo_template.pptx`)写 outline(donor + slots + drop + fit),然后:

```bash
python scripts/build_deck.py build "<模板路径>" outline.json 输出.pptx
python scripts/build_deck.py verify 输出.pptx        # 占位残留必须清零
```

### 第 5 步:预览校验(不可省略)

```bash
python scripts/export_preview.py 输出.pptx 预览目录
```
逐页看图,按 `references/页面版式模式.md` 第 4 节清单查:文字裁切/品牌残留/目录对应/构图重复/字号下限/数据口径。
**发现问题只改 outline 重跑,不手改 pptx**(手改会在下次重跑时丢失)。重复 4–5 步直到全过。

## 装配纪律(违反即返工)

- **字数对齐**:新文本字符数 ≈ 原文本(汉字按 2、拉丁按 1 折算);≥60 pt 槽位中文 ≤4 字,≥100 pt ≤3 字;封面主标宁短勿长
- **品牌清理**:模板自带 logo(右上角无字组合)、版权页、论坛/商城地址、设计师署名,一律 drop 或弃用该页
- **目录/章节一一对应**,条目 ≤10 字;正文 ≥12 pt,单行 ≤30 字
- **fit 是兜底不是万能**:大字号槽位超长仍会裁切,以预览为准
- **数据口径**:数字带周期与口径;"待补"交付前清零
- 配色只用模板当页已有的色(60-30-10);**业绩数据红=涨、绿=跌**(中国惯例)
- 同一文件多次 Edit/改 outline 串行执行,避免并行写丢

## 目录结构

```
business-report-ppt/
├── SKILL.md                        # 本文件:五步工作流 + 纪律
├── references/
│   ├── 设计系统.md                  # 网格/字体阶梯/11 风格族配色/页面解剖/构图词表
│   ├── 页面版式模式.md              # donor 认页/outline 写法/替换纪律/校验清单/FAQ
│   └── 模板索引.md                  # 91 套同源模板索引 + 风格族推荐表 + 格式警示(BYO 说明见文首)
├── scripts/
│   ├── templates.json              # 模板库索引数据(编号/文件/风格族/明暗/主色板/格式)
│   ├── template_lib.py             # 模板库检索:families/list/pick/path/check(BYO 解析)
│   ├── build_deck.py               # 装配引擎:analyze(提槽位)/build(donor拷贝+替换+drop+fit)/verify
│   ├── export_preview.py           # PowerPoint COM 每页导出 PNG 目检
│   ├── payment_gate.py             # 付费闸口(probe/402 账单/Proof 验签/幂等/履约确认)
│   └── make_demo_template.py       # 生成原创合成演示模板(默认输出到 assets/)
├── assets/
│   └── demo_template.pptx          # 原创合成模板:封面/目录/过渡/卡片/数据/结尾,无模板库时练手用
└── samples/
    ├── demo_outline.json           # 完整示例:8 页《季度业务汇报》(donor 复用演示)
    └── demo_季度汇报.pptx           # 按上述 outline 装配的成品(已预览校验)
```

## 快速命令卡

```bash
pip install python-pptx pywin32                    # 首次(预览需要 PowerPoint)
export BUSINESS_PPT_ROOT="/path/to/你的模板目录"      # 有模板库才需要
python scripts/template_lib.py check               # 模板库状态(未配置时打印 BYO 指引)
python scripts/template_lib.py pick 商务红黑
python scripts/build_deck.py analyze "<模板>" --out slots.json
python scripts/build_deck.py build "<模板>" outline.json out.pptx
python scripts/build_deck.py verify out.pptx
python scripts/export_preview.py out.pptx preview/
python scripts/payment_gate.py gate                  # 付费闸口(402 契约见上文;免费自动放行)
# 无模板库练手:python scripts/make_demo_template.py ../assets/demo_template.pptx
```
