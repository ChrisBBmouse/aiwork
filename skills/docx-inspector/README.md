# docx-inspector

[![CI](https://github.com/ChrisBBmouse/aiwork/actions/workflows/ci.yml/badge.svg)](https://github.com/ChrisBBmouse/aiwork/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](../../LICENSE)
[![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue.svg)](pyproject.toml)

**对任意 .docx 做排版体检。** 不生成文档，只回答一个问题：这份 Word 交出去会不会出问题。

[English](README.en.md)

```bash
python -m docx_inspector 报告.docx
```

```
docx-inspector 0.1.0
文件: 报告.docx  34.8 KB
版心: 14.66 × 24.62 cm   页边距: 左3.17 右3.17 上2.54 下2.54 cm

ERROR  B7   表格总宽 20.00cm，超出可用版心 14.66cm（溢出 5.34cm）。fixed 布局，渲染时必然溢出   第 1 个表格（4 列）

9 项检查 · 1 error / 0 warning / 1 info
结果: 不通过 (exit 1)
```

## 为什么需要它

AI 现在很会"生成" Word 文档。但生成出来的文件，**没有人负责验收**：

| | 能生成 | 能编辑 | 能检查**任意** docx | 能量化判定 |
|---|---|---|---|---|
| Anthropic 官方 docx | ✅ | ✅ | ❌ | ❌ |
| gongwen-gbt9704 | ✅ | ❌ | ❌ 仅自家产物 | 部分 |
| docx-cli | ✅ | ✅ | ❌ | ❌ |
| **docx-inspector** | ❌ | ❌ | ✅ | ✅ |

它不抢生成器的活，它是**所有人的下游**。生成完跑一遍，就知道能不能交。

## 它比"看起来正确"更严格

同一个"表格超宽"，结果完全不同——取决于布局模式：

| 布局 | 声明宽度 | 实际渲染 | 结论 |
|---|---|---|---|
| `auto` | 超 0.38cm | Word **自动压缩**回版心 | ⚠️ warning（但 WPS/在线预览可能真溢出） |
| `fixed` | 超 5.34cm | Word **照画不误**，第 4 列被切在页外 | 🔴 error |

![溢出标注](docs/overflow.png)

`docx-inspector` 区分这两种情况，因为它读的是 `tblLayout`，而不是只看一个数字就报警。

## 渲染级证据：`--annotate`

```bash
python -m docx_inspector 报告.docx --annotate out/
```

它会把文档渲染成图，**在图上用红框标出真正越过版心右边界的区域**，并直接写明超了多少厘米。

坐标不是从 XML 声明里读的，是从渲染结果里量出来的——声明说 14.66cm，渲染才是真相。

## 12 条规则

按 **5 个物理系统**组织。这是刻意的：按"文档场景"分类（论文/审计/合同/标书）永远列不完，但任何 docx 只在这 5 个系统里出问题。

| 系统 | ID | 规则 | 默认 |
|---|---|---|---|
| 包结构 | A1 | ZIP 头非法（含 DLP 透明加密识别） | ✅ |
| | A3 | 关系目标缺失（图片/嵌入对象找不到） | ✅ |
| | A4 | 媒体文件损坏或格式与扩展名不符 | ✅ |
| 页面几何 | B5 | 页边距或版心异常 | ✅ |
| | B7 | 表格总宽超出可用版心（区分 auto/fixed） | ✅ |
| | B8 | 图片显示宽度超出可用版心 | ✅ |
| 字体系统 | C10 | 文档引用的字体本机缺失（会静默替换） | ✅ |
| | C12 | theme 属性覆盖显式字体名（改了不生效） | ⬜ |
| | C14 | 同类元素字体不一致 | ⬜ |
| 流式布局 | D15 | 列宽总和与 tblGrid 声明不符 | ✅ |
| | D17 | 多行表格缺少跨页表头重复标记 | ⬜ |
| 引用编号 | E20 | 编号 numId 跨批复用（列表串号） | ✅ |

默认只开 9 条。其余按需开启——**规则全开会淹没用户，最后被关掉**，这是 linter 最常见的死法。

```bash
python -m docx_inspector f.docx --preset thesis   # 论文/长报告：+C12 +C14 +D17
python -m docx_inspector f.docx --preset blockers # 只看硬伤：A1 A3 A4 B7 B8
python -m docx_inspector f.docx --only B7,E20     # 只跑指定规则
```

## 嵌进你的流程

```bash
# CI / agent 链：--json + 退出码
python -m docx_inspector out.docx --json > report.json
echo $?   # 0 通过 / 1 有 error / 2 无法解析
```

```json
{
  "tool": "docx-inspector",
  "counts": {"error": 1, "warning": 0, "info": 1},
  "findings": [
    {
      "rule": "B7",
      "severity": "error",
      "message": "表格总宽 20.00cm，超出可用版心 14.66cm（溢出 5.34cm）。fixed 布局，渲染时必然溢出",
      "location": "第 1 个表格（4 列）",
      "fixable": true,
      "detail": {"table_index": 0, "overflow_cm": 5.34, "layout": "fixed"}
    }
  ]
}
```

如果你在写文档生成的 skill：**生成完之后调一下这个，再交付**。

## 安装

**作为 Agent Skill（Claude Code / Codex / OpenCode / Chatbox 等）：**

```bash
npx skills add ChrisBBmouse/docx-inspector
```

`SKILL.md` 在仓库根目录，安装时无需 `--skill` 参数。

**作为命令行工具：**

```bash
pip install git+https://github.com/ChrisBBmouse/docx-inspector
# 或克隆后 pip install -e .
docx-inspector 报告.docx
```

**只想要本体：**

```bash
git clone https://github.com/ChrisBBmouse/docx-inspector.git
cd docx-inspector
pip install python-docx
python -m docx_inspector 报告.docx
```

## 依赖

```bash
pip install python-docx
# 可选：渲染标注图需要
pip install pymupdf pillow
# Windows 有 Word 即可渲染；其他平台需要 LibreOffice（soffice）
```

无第三方 CLI 框架，纯 argparse。

## 定位与搭配建议

本 skill 只做一件事：检查。推荐与文档生成类 skill 搭配使用：

```
用任意 docx skill 生成  →  docx-inspector 检查  →  按报告修正  →  交付
```

这样能把「生成」和「验收」分开：生成器负责产出，检查器负责保证规范性，
两边都不必为对方的功能负责。

- **不是生成器** —— 文档创建交给专门的 doc/文档 skill（Anthropic `docx`、`docx-cli`、`gongwen-gbt9704` 等），先生成、再检查
- **不做内容评价** —— 措辞、逻辑、完整性没有唯一正解
- **不做业务规范引擎** —— 审计勾稽、期刊格式属于数据，不是代码
- **规则永远不超过 30 条** —— 加一条必须先减一条

## License

MIT
