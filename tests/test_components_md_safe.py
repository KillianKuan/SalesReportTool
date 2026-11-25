import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from components import md_safe  # noqa: E402


class MdSafeTests(unittest.TestCase):
    def test_escapes_dollar_signs(self):
        self.assertEqual(md_safe("NT$1.39B vs NT$2.89B"), r"NT\$1.39B vs NT\$2.89B")

    def test_already_escaped_is_left_alone(self):
        self.assertEqual(md_safe(r"NT\$5"), r"NT\$5")

    def test_none_and_non_string(self):
        self.assertEqual(md_safe(None), "")
        self.assertEqual(md_safe(12), "12")

    def test_text_without_dollar_unchanged(self):
        self.assertEqual(md_safe("營收年減 46%"), "營收年減 46%")


if __name__ == "__main__":
    unittest.main()
