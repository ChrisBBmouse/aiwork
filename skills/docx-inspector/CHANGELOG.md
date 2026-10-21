# Changelog

格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，版本号遵循语义化版本。

## [0.1.2] - 2026-10-21

本版由**两轮独立的外部测试**驱动（WSL2 环境，共 15 份真实业务文档），
修复 3 个缺陷、补齐 4 处测试缺口。

### Fixed

- **C10 在非 Windows 平台上完全失效** —— `installed_fonts()` 只读 Windows 注册表与
  `C:\Windows\Fonts`，在 Linux / macOS 上返回空集，C10 退化为「无法枚举本机字体，
  跳过检查」，字体检查实际处于关闭状态。
- **页面属性缺失导致整篇无法解析**（静默漏检，危害高于误报）—— 部分文档的
  `sectPr` 只声明 `pgSz` 而不声明 `pgMar`，python-docx 对缺失项返回 `None`，
  `Emu(None)` 抛 `TypeError`，被上层当成「无法解析」整篇跳过。现在按 OOXML
  规范默认值（A4 / 四边 1440 twips）推断，并由 `B5` 明确告知「版心为推断值」。
- **D15 对横向合并单元格重复计数** —— `row.cells` 会把 `gridSpan` 合并单元格按
  逻辑列展开，同一个 `tcW` 被累加多次（实测一个跨 4 列的表头把 14.62cm 累加成
  58.49cm）。改为遍历原始 `<w:tc>`，并兼容内容控件包裹的单元格。

### Added

- 字体枚举扩展为四路探测：Windows 注册表 → Windows 字体目录 → fontconfig
  （`fc-list`，Linux）→ macOS（`system_profiler`）。
- 枚举结果用 `functools.lru_cache` 缓存，同一进程内只探测一次
  （Windows 上 273 个字体族耗时 4 ms）。
- `tests/test_fonts.py`：字体枚举的单元测试（fc-list / system_profiler / 缓存）。
  这两条路径在 Windows 上无法端到端验证，只能靠 mock 锁住。
- `case_no_pgmar.docx` / `case_merged_header.docx` 两个 fixture，以及对应的回归
  断言与「推断值必须可见」的专项检查。回归用例 5 → 7。
- CI 的 ubuntu job 现在能真正执行 C10 断言（此前枚举为空会自动跳过），
  并新增单元测试步骤。

### Note

> **「跳过」不等于「通过」。** 此前 CI 的 4/4 全绿掩盖了 C10 的实际失效：
> ubuntu 上 `HAS_FONT_ENV` 为 `False`，测试自动 SKIP 了断言。
>
> 同理，`B5` 的推断提示是必须的——修掉旧的静默漏检，不能引入新的静默假设。
> 版心一旦是推断出来的，就必须告诉用户「结论仅供参考」。

### Credit

三个缺陷均由独立的外部交叉测试发现（WSL2 / Python 3.12 / 15 份真实业务文档，
两轮独立测试）。缺陷报告附带完整证据链与回滚验证，修复方案在此基础上补充了
「推断值可见」与「原始 tc 遍历」两处改进。

## [0.1.1] - 2026-10-21

用 5 份真实业务文档做压测后，修复了误报与噪音。**默认规则现在在这批文档上零误报。**

### Fixed

- **C12 误报**：原先扫描全部样式，包括模板带来的、从未被引用的内置样式
  （`macro` 用 Courier、`HTML Preformatted` 等），导致 5/5 文档全部误报
  "theme 字体覆盖显式字体"。现在只检查文档实际引用到的样式。
- **D15 重复报告**：同一文档里多个表格存在同类不一致时，原先每个表格报一条
  （实测一份文档刷 3 条完全相同的行）。现合并为一条，`detail.tables` 给出全部表格索引。
- **C10 回归**：重构 `used_style_names` 时漏掉 `doc = self.doc`，触发 `NameError`。
  由 fixture 回归测试捕获并修正。

### Added

- `tests/manual_annotate.py`：渲染标注链路（`--annotate`）的手动验证脚本。
  该链路依赖 Word COM / LibreOffice，无法在 GitHub runner 上跑，故不进 CI，
  但保留脚本以便改动后自查。
- 本 CHANGELOG。

### 压测结论

在一批实际业务文档（5 份，含模板）上：

| 结果 | 数量 | 说明 |
|---|---|---|
| 完全 clean | 3 份 | 零 findings |
| 仅 B7（表格声明超版心） | 2 份 | 真阳性：声明宽度 15.49cm / 15.04cm，版心 14.66cm |
| 另有 D15（tblGrid 与单元格列宽不一致） | 1 份 | 真阳性：3 个表格 |

**误报 0 条。** 另有一条 `C14`（正文字体不一致）在 `--preset thesis` 下命中，
经核实为真阳性——正文后段粘入了使用不同字体的外部内容。

> 教训：自建的 5 个 fixture 全部通过，但真实文档一跑就 100% 触发误报。
> **fixture 测不出真实数据的脏。**

## [0.1.0] - 2026-10-21

首个版本。

### Added

- 12 条规则，按 5 个物理系统组织（包结构 / 页面几何 / 字体 / 流式布局 / 引用编号），
  默认开启 9 条
- `B7` 按 `tblLayout` 分档：`fixed` → error（必然溢出），`auto` → warning
- `--json` 稳定 schema + 语义化退出码（0 通过 / 1 有 error / 2 无法解析）
- `--annotate` 渲染标注图，坐标从渲染后的 PDF 实测，而非读取 XML 声明
- `--preset plain|thesis|blockers`、`--only`、`--skip`
- 5 个 fixture 回归测试，CI 覆盖 ubuntu/windows × Python 3.9/3.12
