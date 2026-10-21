# aiwork

A workbench for AI tools and agent skills. Every item follows one rule: **verifiable, bounded, no hype**.

[中文说明](README.md)

## Skills

| Skill | What it does | Install |
|---|---|---|
| [docx-inspector](skills/docx-inspector/) | Audit any `.docx` for layout integrity — 12 rules across package structure, page geometry, fonts, flow layout and numbering | `npx skills add ChrisBBmouse/aiwork --skill docx-inspector` |

## Layout

```
aiwork/
├── skills/
│   └── <skill-name>/
│       ├── SKILL.md          # required, must sit at the skill root
│       ├── README.md         # human-facing docs
│       ├── pyproject.toml    # pip install -e
│       ├── <package>/        # implementation
│       └── tests/            # standalone runnable
├── tools/                    # standalone tools (non-skill form)
├── .github/workflows/ci.yml  # discovers and tests every skill
├── LICENSE
└── README.md
```

**Adding a skill** only means: create a directory under `skills/`, add `SKILL.md` and
`tests/run_tests.py`. CI discovers and runs it automatically — no workflow edit needed.

## Design principles

1. **Deterministic checks only** — no opinions on wording, logic or aesthetics; those have no single right answer.
2. **Minimal default set** — enabling every rule drowns the user, and they turn the linter off. That is how most linters die.
3. **Fewer false positives beats more coverage** — better to miss something than to make users distrust the tool.
4. **Say what we don't do** — capability boundaries live in `SKILL.md`, not left for users to discover by trial and error.

## License

MIT
