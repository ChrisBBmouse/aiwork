# aiwork
2026/10/21-new v0.0.1

一个放 AI 工具与 Agent Skill 的工作台。每个子项的规则是：**能验证、有边界、不吹**。

[English](README.en.md)

## Skills

| Skill | 做什么 | 安装 |
|---|---|---|
| [docx-inspector](skills/docx-inspector/) | 对任意 `.docx` 做排版体检：12 条规则覆盖包结构、页面几何、字体、布局、编号 | `npx skills add ChrisBBmouse/aiwork --skill docx-inspector` |

## 目录约定

```
aiwork/
├── skills/
│   └── <skill-name>/
│       ├── SKILL.md          # 必需，且必须在本目录根
│       ├── README.md         # 给人看的说明
│       ├── pyproject.toml    # 可 pip install -e
│       ├── <package>/        # 实现
│       └── tests/            # 可独立运行
├── tools/                    # 独立工具（非 skill 形态）
├── .github/workflows/ci.yml  # 统一跑所有 skills 的测试
├── LICENSE
└── README.md
```

**新增一个 skill 只需要**：在 `skills/` 下建目录，放好 `SKILL.md` + `tests/run_tests.py`。CI 会自动发现并跑它，不需要改 workflow。

## 设计原则

1. **只做确定性判断** —— 不评价内容好坏、措辞、审美，那些没有唯一正解
2. **默认最小集** —— 规则全开会淹没用户，最后被关掉
3. **误报率 > 覆盖率** —— 宁可漏报，也不让用户怀疑工具
4. **明确写不做什么** —— 能力边界写在 SKILL.md 里，不靠用户试出来

## License

MIT
