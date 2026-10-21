---
name: docx-inspector
description: "对任意 .docx 做排版体检，不生成文档。检查包结构损坏、表格/图片超出可用版心、字体本机缺失、theme 覆盖字体、列宽与 tblGrid 不符、编号跨批串号等 12 项确定性问题，可选渲染成图并用红框标出实际溢出区域。触发词：检查Word、文档排版有问题、表格超宽、Word打开异常、字体变了、编号乱了、交付前验收、docx 体检、linter。亦触发：audit docx、check word layout、table too wide、table overflow、font missing、numbering restart、word file broken、docx linter、validate word document。"
category: knowledge
version: 0.1.2
author: 本机自研
license: MIT
---

# docx-inspector

**只做下游**：任何工具生成的 .docx 都能拿来检查。不生成、不编辑、不评价内容。

> English docs: [README.en.md](README.en.md)
>
> 下文所有 `python -m docx_inspector` 命令，**均在本 skill 目录下执行**
> （即 SKILL.md 所在目录，含 `docx_inspector/` 包）。
> 若在其他目录，可用 `set PYTHONPATH=<skill目录>` 或 `cd` 过去。

## 何时触发

- 用户说"这份 Word 有问题/打不开/表格超宽/字体变了/编号乱了"
- 用户要交付文档，想先验收
- **你自己刚生成了一份 .docx，交付前想确认没硬伤**（最常见）
- 需要判断某份文档在 WPS / 在线预览里会不会变形

不触发：要从零生成文档（那是生成器的事）；只改内容不改格式。

## 用法

```bash
# 零配置体检，人类可读
python -m docx_inspector 文件.docx

# 机器可读（推荐嵌进流程）
python -m docx_inspector 文件.docx --json

# 渲染成图并标红框（需要 Word COM 或 LibreOffice）
python -m docx_inspector 文件.docx --annotate out/

# 场景 preset
python -m docx_inspector 文件.docx --preset thesis     # 论文：+C12 +C14 +D17
python -m docx_inspector 文件.docx --preset blockers   # 只看硬伤
python -m docx_inspector 文件.docx --only B7,E20       # 只跑指定规则
python -m docx_inspector 文件.docx --skip C10          # 跳过指定规则
```

## 输出契约（嵌进 agent 链时看这里）

**退出码**

| 码 | 含义 | agent 该怎么做 |
|---|---|---|
| 0 | 通过（可能仍有 warning/info） | 可交付 |
| 1 | 存在 error | 修完再交付 |
| 2 | 无法解析该文件 | 别重试，报告给用户 |

**JSON 结构**

```json
{
  "tool": "docx-inspector",
  "counts": {"error": 1, "warning": 0, "info": 1},
  "rules_run": ["A1", "..."],
  "rules_skipped": ["C12"],
  "findings": [{
    "rule": "B7", "severity": "error",
    "message": "...", "location": "第 1 个表格（4 列）",
    "fixable": true,
    "detail": {"table_index": 0, "overflow_cm": 5.34, "layout": "fixed"}
  }]
}
```

`detail` 里的字段是稳定的，可直接用于后续自动修复。

## 定位：下游检查器，建议搭配使用

**如果你在写文档生成的 skill：生成完之后调这个，再交付。**

它不生成、不编辑，所以不会跟生成器冲突。典型接法：

```bash
python -m docx_inspector out.docx --json || echo "有 error，需要修"
```

## 规则与系统

按 5 个物理系统组织，不按文档场景（场景无穷，物理失效只有 5 类）：

| 系统 | 规则 |
|---|---|
| A 包结构 | A1 ZIP头/加密、A3 关系目标缺失、A4 媒体损坏 |
| B 页面几何 | B5 边距异常、B7 表格超版心、B8 图片超版心 |
| C 字体系统 | C10 字体本机缺失、C12 theme覆盖字体、C14 同元素字体不一致 |
| D 流式布局 | D15 列宽与tblGrid不符、D17 缺跨页表头 |
| E 引用编号 | E20 编号跨批复用 |

## 判定要点（避免误读）

1. **B7 分两档**：`fixed` 布局→error（必然溢出）；`auto` 布局→warning（Word 会压缩，但 WPS/在线预览可能真溢出）。不要看到 warning 就说"没问题"。
2. **C10 是 warning 不是 error**：字体缺失时 Word 会静默替换，文档能打开，但字形变了。若要严格交付，把它当 error 处理。
3. **B5 的 info 是"报平安"**，不是问题。
4. **A1 能识别 DLP 透明加密**（文件头 `E0 A8 91 E7`）——这类文件普通程序读到的是密文，必须先解密或重新明文落盘。

## 能力边界

**能查**：包结构、几何溢出、字体存在性、布局一致性、编号复用。

**不能查**：
- 内容质量（措辞、逻辑、完整性）
- 业务规范（审计勾稽、期刊格式、公文国标）
- 审美判断
- 脚注/域代码是否正确计算（Word 打开时才计算）

以上需要基准的判定，用 `--profile`（v1.1 计划）或外部 spec，不进内核。

## 依赖

```bash
pip install python-docx          # 必需
pip install pymupdf pillow       # 仅 --annotate 需要
# 渲染后端：Windows 用 Word COM；其他平台需 LibreOffice (soffice)
```

渲染后端不可用时 `--annotate` 会报错但不影响检测结果。
