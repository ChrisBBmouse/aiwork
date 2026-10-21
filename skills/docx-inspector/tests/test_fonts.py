# -*- coding: utf-8 -*-
"""字体枚举的单元测试

为什么需要：`_fonts_from_fontconfig()` / `_fonts_from_macos()` 在 Windows 上无法
端到端跑通，只能靠 mock 锁住解析逻辑。CI 的 ubuntu job 另有一层真实环境验证。

运行：python tests/test_fonts.py
"""
import os
import sys
import unittest
from unittest import mock

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
sys.path.insert(0, ROOT)

from docx_inspector import rules                          # noqa: E402


class FontConfigParsing(unittest.TestCase):
    """fc-list 解析"""

    def test_splits_comma_separated_families(self):
        fake = ("Noto Sans CJK SC,Noto Sans CJK SC Regular\n"
                "DejaVu Sans,DejaVu Sans Book\n")
        with mock.patch.object(rules.shutil, "which", return_value="/usr/bin/fc-list"):
            with mock.patch.object(rules.subprocess, "run",
                                   return_value=mock.Mock(stdout=fake)):
                got = rules._fonts_from_fontconfig()
        self.assertIn("Noto Sans CJK SC", got)
        self.assertIn("Noto Sans CJK SC Regular", got)
        self.assertIn("DejaVu Sans", got)
        self.assertIn("DejaVu Sans Book", got)

    def test_returns_empty_when_fc_list_absent(self):
        with mock.patch.object(rules.shutil, "which", return_value=None):
            self.assertEqual(rules._fonts_from_fontconfig(), set())

    def test_returns_empty_on_subprocess_exception(self):
        with mock.patch.object(rules.shutil, "which", return_value="/usr/bin/fc-list"):
            with mock.patch.object(rules.subprocess, "run",
                                   side_effect=OSError("boom")):
                self.assertEqual(rules._fonts_from_fontconfig(), set())

    def test_ignores_blank_and_whitespace(self):
        fake = "  ,Arial,\n  , ,\n"
        with mock.patch.object(rules.shutil, "which", return_value="/usr/bin/fc-list"):
            with mock.patch.object(rules.subprocess, "run",
                                   return_value=mock.Mock(stdout=fake)):
                got = rules._fonts_from_fontconfig()
        self.assertEqual(got, {"Arial"})


class MacOsParsing(unittest.TestCase):
    """system_profiler 解析"""

    def test_extracts_family_and_full_name(self):
        fake = (
            "Fonts:\n"
            "      Family: Helvetica\n"
            "      Full Name: Helvetica\n"
            "      Family: PingFang SC\n"
            "      Full Name: PingFang SC Regular\n"
        )
        with mock.patch.object(rules.shutil, "which", return_value="/usr/sbin/system_profiler"):
            with mock.patch.object(rules.subprocess, "run",
                                   return_value=mock.Mock(stdout=fake)):
                got = rules._fonts_from_macos()
        self.assertIn("Helvetica", got)
        self.assertIn("PingFang SC", got)
        self.assertIn("PingFang SC Regular", got)

    def test_returns_empty_when_binary_absent(self):
        with mock.patch.object(rules.shutil, "which", return_value=None):
            self.assertEqual(rules._fonts_from_macos(), set())


class Caching(unittest.TestCase):
    """枚举结果必须缓存，避免每次检查都扫一遍字体"""

    def setUp(self):
        rules._installed_fonts_cached.cache_clear()

    def tearDown(self):
        rules._installed_fonts_cached.cache_clear()

    def test_returns_identical_object(self):
        """两次调用返回同一个 frozenset 对象 → 说明走了缓存"""
        a = rules._installed_fonts_cached()
        b = rules._installed_fonts_cached()
        self.assertIs(a, b)

    def test_second_call_is_cache_hit(self):
        rules.installed_fonts()          # miss
        rules.installed_fonts()          # hit
        info = rules._installed_fonts_cached.cache_info()
        self.assertGreaterEqual(info.hits, 1)
        self.assertEqual(info.maxsize, 1)


class PublicApi(unittest.TestCase):
    def test_returns_mutable_set(self):
        got = rules.installed_fonts()
        self.assertIsInstance(got, set)
        got.add("__probe__")             # 不应污染缓存的 frozenset

    def test_pseudo_fonts_recognised(self):
        for name in ("", "default", "inherit"):
            self.assertIn(name.strip().lower(), rules._PSEUDO_FONTS)


if __name__ == "__main__":
    unittest.main(verbosity=2)
