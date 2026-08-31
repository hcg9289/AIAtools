import importlib.util
import os
import tempfile
import time
import unittest
from pathlib import Path


MODULE_PATH = Path(__file__).parents[1] / "ops" / "runtime_retention.py"
SPEC = importlib.util.spec_from_file_location("runtime_retention", MODULE_PATH)
runtime_retention = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(runtime_retention)


class RuntimeRetentionTests(unittest.TestCase):
    def test_dry_run_counts_without_removing(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir).resolve()
            runtime_retention.ALLOWED_ROOTS = frozenset({root})
            old_file = root / "old.pdf"
            old_file.write_bytes(b"old")
            os.utime(old_file, (1, 1))

            count, size = runtime_retention.cleanup_root(
                root, time.time() - 60, dry_run=True
            )

            self.assertEqual((count, size), (1, 3))
            self.assertTrue(old_file.exists())

    def test_cleanup_removes_only_old_regular_files(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir).resolve()
            runtime_retention.ALLOWED_ROOTS = frozenset({root})
            old_file = root / "old.pdf"
            fresh_file = root / "fresh.pdf"
            keep_file = root / ".gitkeep"
            old_file.write_bytes(b"old")
            fresh_file.write_bytes(b"fresh")
            keep_file.write_bytes(b"")
            os.utime(old_file, (1, 1))
            os.utime(keep_file, (1, 1))

            count, size = runtime_retention.cleanup_root(
                root, time.time() - 60, dry_run=False
            )

            self.assertEqual((count, size), (1, 3))
            self.assertFalse(old_file.exists())
            self.assertTrue(fresh_file.exists())
            self.assertTrue(keep_file.exists())

    def test_rejects_unapproved_root(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir).resolve()
            runtime_retention.ALLOWED_ROOTS = frozenset()
            with self.assertRaises(ValueError):
                runtime_retention.cleanup_root(root, time.time(), dry_run=True)


if __name__ == "__main__":
    unittest.main()
