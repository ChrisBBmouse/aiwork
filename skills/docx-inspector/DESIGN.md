# docx-inspector：研究与实施计划

> 2026-10-21 · 目标：**开源发布，积累个人影响力（方案 B）**
> 状态：v1.0 开发中
> 本文件是项目的唯一设计依据，改方向先改这里。

---

## 一、调研结论（决定方向的三条事实）

### 事实 1：两条显而易见的赛道都已饱和

| 赛道 | 现状 | 代表 |
|---|---|---|
| 通用 docx 生成 | 红海 | `anthropics/skills/docx`（生产在用，仓库 179.7K star）、Nous `hermes-agent/docx`、`kklimuk/docx-cli` |
| 中文公文 GB/T 9704 | **红海，13+ 同类**，最高 700 星 | `mizzlelover/gongwen-gbt9704-skill`、`Rimagination/gongwen-draft`、ClawHub/hunkguo 等 |
| 模板格式复刻 | 有实现，且有付费产品 | `iamzhihuix/docx-format-replicator`、`fokkyp/docx-toolkit` Scenario C、formatmagic.ai（$29.95/月） |

**结论：不做生成器，不做公文排版，不做格式复刻。**

### 事实 2：现有工具的验收都是"绑定自己产物"的

- `gongwen-gbt9704-skill` 有 `verify_gongwen_docx.mjs`，但只校验它自己生成的 profile
- Anthropic 官方有渲染验证（soffice → pdftoppm → 看图），但无量化判据
- `hermes-agent` 有 `docx_validate.py`，但只是包健康检查，非排版判断
- **没有一个工具能对任意 .docx 做排版体检**

### 事实 3：微软自己都在收缩规则集

Word 内置 Accessibility Checker 按**问题类别**组织（不按文档类型），且 Word 365 更新后**砍掉了 4 项检查**（Repeated Blank Characters / Infrequent Headings / Table Header Rows / Unclear Hyperlink Text）。

**结论：臃肿不是"做得不够多"，是"做多了必然打折"。规则集必须克制。**

---

## 二、定位

> **docx-inspector：读懂一份 Word 文档的排版真相——提取规范、诊断问题、验收产出。**

三条定位原则：

1. **只做下游，不做上游**。任何人生成的 docx 都能拿来检查。不抢生成器的活，因此不树敌。
2. **只做确定性判断**。不评价内容好坏、措辞优劣、结构是否合理——那些没有唯一正解。
3. **默认零配置**。`docx-inspector file.docx` 直接出结果，不需要先理解我们的规范体系。

### 与竞品的能力边界

| 能力 | Anthropic docx | gongwen | docx-cli | **docx-inspector** |
|---|---|---|---|---|
| 生成文档 | ✅ | ✅ | ✅ | ❌ 不做 |
| 公文规范 | ❌ | ✅ | ❌ | ❌ 不做 |
| 编辑/redline | ✅ | ❌ | ✅ | ❌ 不做 |
| **任意 docx 体检** | ❌ | ❌（仅自家产物） | ❌ | ✅ |
| **量化判据 + 退出码** | ❌ | 部分 | ❌ | ✅ |
| **可视化标注图** | ❌ | 有截图证据链 | ❌ | ✅ |

---

## 三、架构：5 个物理系统

**核心判断：按"文档场景"分类无穷无尽，按"失效物理机制"分类只有 5 类且已穷尽。**

科研报告的"表格超宽"和审计报告的"表格超宽"是同一个 bug。

| 系统 | 覆盖故障 | 与场景相关 |
|---|---|---|
| **A 包结构** | 打不开、提示修复、图片丢失 | ❌ |
| **B 页面几何** | 表格/图片超版心、页边距异常、页脚越界 | ❌ |
| **C 字体系统** | 豆腐块、静默替换、中英字体冲突、theme 覆盖 | ❌ |
| **D 流式布局** | 表头跨页丢失、列宽错位、单元格越界 | ❌ |
| **E 引用与编号** | 编号串号、交叉引用失效、目录不同步 | ❌ |

### 12 条 v1.0 规则

| ID | 系统 | 规则 | 默认 | 可图示 |
|---|---|---|---|---|
| A1 | 包结构 | ZIP 头非法（含 DLP 加密识别） | ✅ | ⭐ |
| A3 | 包结构 | 关系目标缺失（图片/嵌入对象找不到） | ✅ | ⭐ |
| A4 | 包结构 | 媒体文件损坏或格式不符 | ✅ | ⭐ |
| B5 | 页面几何 | 页边距异常（负数/过大/版心过窄） | ✅ | ⭐⭐ |
| B7 | 页面几何 | **表格总宽超出可用版心**（auto→warning / fixed→error） | ✅ | ⭐⭐⭐ |
| B8 | 页面几何 | **图片显示宽度超出可用版心** | ✅ | ⭐⭐⭐ |
| C10 | 字体系统 | **文档引用的字体本机缺失** | ✅ | ⭐⭐⭐ |
| C12 | 字体系统 | theme 属性覆盖显式字体（静默变样） | ⬜ | ⭐⭐ |
| C14 | 字体系统 | 同类元素字体不一致（正文漂移） | ⬜ | ⭐⭐ |
| D15 | 流式布局 | 列宽总和与 tblGrid 声明不符 | ✅ | ⭐⭐ |
| D17 | 流式布局 | 跨页表格缺少表头重复标记 | ⬜ | ⭐⭐ |
| E20 | 引用编号 | **编号 numId 跨批复用（串号）** | ✅ | ⭐⭐⭐ |

默认开启 9 条（✅），3 条 opt-in（⬜）。

> **实现中发现的修正（2026-10-21）**
> B7 原本定为 error。实测发现 `auto` 布局的表格在 Word 渲染时会被自动压缩回版心
> （真实模板声明超 0.38cm，渲染未溢出），此时报 error 就是误报。
> 改为按 `tblLayout` 分档：`fixed` → error（必然溢出，实测 20cm 表格第 4 列被切在页外）；
> `auto` → warning（Word 会救，但 WPS / 在线预览可能不救）。
> 这是「误报率 > 覆盖率」铁律的第一次实际应用。

### 为什么"代码量是常数"

```
场景数 N   ↑ 无限增长（论文/审计/合同/标书/…）
代码量     =  常数（12~25 条规则）
配置文件   ↑ 每个场景 5 行 preset
```

Preset 只是规则清单，不含新代码：

```yaml
# presets/thesis.yaml
extra_on:  [C12, C14, D17]
extra_off: [C10]        # 学术文档字体环境固定
```

---

## 四、防臃肿四铁律

来自 docsie/Fern 的 linting 最佳实践：
> "在存量语料上开启全部规则会立刻产生数千个违规，把作者淹没，最终导致团队**直接关掉这个 linter**。"

因此：

1. **默认最小集** —— 默认只开 9 条最致命的，其余 opt-in
2. **准入双门槛** —— 必须有①确定性判据 ②明确修复路径或标注 `unfixable`
3. **误报率 > 覆盖率** —— 单条规则误报率 >5% 直接下线，宁可漏报
4. **硬上限 30 条** —— 超了必须先删一条再加

---

## 五、v1.0 范围（本次）

### 做

```
12 条检测规则，默认开 9 条
├ 重点打磨 4 条（可图示）：B7 / B8 / C10 / E20
│   └ 每条产出"标注图"：在渲染图上画红框/竖线
├ --json 稳定 schema + 退出码（0 通过 / 1 有 error / 2 无法解析）
├ --preset plain | thesis
├ --annotate 输出标注 PNG
├ 3 个测试用例（正常 / 表格溢出 / 字体漂移）
└ README 首屏 = 4 张标注图 + 一句话定位
```

### 不做

- ❌ 生成文档（一做就变竞品）
- ❌ 内容检查（措辞/逻辑/完整性）
- ❌ 业务规范（勾稽关系/期刊格式）—— 那是 spec，不是代码
- ❌ 规则市场/插件系统
- ❌ GUI
- ❌ 超过 30 条规则

---

## 六、路线图

```
v1.0（本次）  12 条检测 + JSON + annotate + 2 preset
v1.1          --fix 两条（B7 表格宽、C14 字体统一）
              理由：fix 要改文件，首次发布不承担毁文档风险；
                    分两次发比一次发大更利于持续曝光
v1.2          --against 基准文档 diff  ← 差异化卖点，最适合 DEMO
之后          传播动作（见下）
```

### 传播动作（v1.2 后）

1. GitHub 仓库 + README（首屏 4 张标注图）
2. ClawHub 发布：
   ```bash
   npm i -g clawhub && clawhub login
   clawhub skill publish ./docx-inspector --slug docx-inspector \
     --categories knowledge --topics "docx,word,validation,quality"
   ```
   硬性要求：GitHub 账号 ≥1 周；frontmatter 单行 key；VirusTotal 扫描；首版不进搜索（pending）
3. skills.sh：`npx skills add <owner>/docx-inspector`
4. 在 Anthropic docx / hermes docx 的 issue 区礼貌提"生成后可用此验收"

---

## 七、关键技术决策

| 决策 | 选择 | 理由 |
|---|---|---|
| 语言 | Python | python-docx / PyMuPDF 已就绪 |
| 渲染后端 | Word COM → LibreOffice → 跳过 | Windows 原生优先，跨平台降级 |
| 渲染 → 图 | Word 导出 PDF → PyMuPDF → PNG | 已验证链路 |
| 标注坐标来源 | **PDF 渲染结果**，不是 OOXML 声明 | 不信声明，信渲染（gongwen 坐标量测同思路） |
| 打包 | 纯脚本 + argparse，无第三方 CLI 框架 | 依赖越少越好分发 |
| 分发形态 | `npx` / `pip` / 直接 python | 便于嵌进别人的 agent 链 |

### 已知坑（来自既有实践）

- theme 字体属性（`asciiTheme`/`eastAsiaTheme`）会覆盖显式字体名
- 每批列表需独立 `abstractNum`+`numId` 才重新编号
- 项目符号须用 U+2022，`U+F0B7` 需 Symbol 字体
- 段落级 `ind` 会覆盖编号定义的悬挂缩进
- 字体名有中英双名：`SimSun` 查不到、`宋体` 有
- **仿宋_GB2312 在部分环境不存在**（国标公文要求的字体）
- PowerShell 内联 Python 遇中文/正则引号会崩，一律写 .py 文件

---

## 八、参考与依据

| 来源 | 借鉴点 |
|---|---|
| `gongwen-gbt9704-skill` | 逐条执行矩阵（自动/校验/人工三档）、截图证据链、坐标量测、字体缺失显式失败 |
| `kklimuk/docx-cli` | 能力做成 CLI、SKILL.md 只教定位模型、弱模型基准测试、退出即正确 |
| `anthropics/skills/docx` | "模型懂 API，这里只讲 footguns"、安全步骤、数据化 gotchas |
| `nousresearch/hermes-agent/docx` | 多 CLI 分工 + JSON 输出 + Quick Reference 表 |
| Microsoft Accessibility Checker | 按问题类别而非文档类型分类；Error/Warning/Tip 三级；**主动收缩规则集** |
| USPTO DOCX 检查器 | 垂直场景官方检查器也只收敛到物理层 |
| docsie/Fern linting 指南 | 最小规则集起步、误报淹没用户、增量扩规则 |
