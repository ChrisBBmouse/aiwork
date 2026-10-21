# -*- coding: utf-8 -*-
"""docx-inspector — 对任意 .docx 做排版体检

设计依据见仓库 DESIGN.md。核心立场：
  1. 只做下游（检查别人的产物），不做生成
  2. 只做确定性判断，不评价内容与审美
  3. 默认零配置
"""
from . import core          # noqa: F401
from . import rules         # noqa: F401  触发规则注册

__version__ = "0.1.1"
__all__ = ["core", "rules"]
