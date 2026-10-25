import os
import unittest
from tempfile import TemporaryDirectory

from qt_ui import ProcessingThread


class TestProcessingThread(unittest.TestCase):
    """Drives the worker logic without a GUI (no QApplication needed for these methods)."""

    def _two_files_with_same_name(self, root):
        paths = []
        for sub, content in (("A", b"%PDF-first"), ("B", b"%PDF-second")):
            os.makedirs(os.path.join(root, sub))
            p = os.path.join(root, sub, "inv.pdf")
            with open(p, "wb") as f:
                f.write(content)
            paths.append(p)
        return paths

    def test_same_basename_files_are_both_copied(self):
        with TemporaryDirectory() as root, TemporaryDirectory() as temp_dir:
            thread = ProcessingThread(self._two_files_with_same_name(root), 0.6)
            copied = thread.copy_files_to_temp(temp_dir)
            self.assertEqual(len(set(copied)), 2)
            contents = {open(p, "rb").read() for p in copied}
            self.assertEqual(contents, {b"%PDF-first", b"%PDF-second"})

    def test_problems_are_returned_with_original_file_names(self):
        """Unreadable PDFs end up in the problems list, named as the user selected them, not by temp name."""
        with TemporaryDirectory() as root:
            thread = ProcessingThread(self._two_files_with_same_name(root), 0.6)
            output_path, df, problems = thread.process_invoices()
            self.assertTrue(output_path and os.path.exists(output_path))
            os.remove(output_path)
            self.assertEqual(len(problems), 2)
            # two selected files share a name, so the folder is shown too
            self.assertTrue(any("A/inv.pdf" in m for m in problems) and any("B/inv.pdf" in m for m in problems))
            for msg in problems:
                self.assertIn("inv.pdf", msg)
                self.assertNotIn("0000_", msg)
                self.assertNotIn("0001_", msg)


class TestUniquePaths(unittest.TestCase):
    def test_same_file_selected_twice_is_kept_once(self):
        from qt_ui import unique_paths
        paths = [r"C:\Facturi\a.pdf", r"C:\FACTURI\A.PDF", r"C:\Facturi\..\Facturi\a.pdf", r"C:\Facturi\b.pdf"]
        self.assertEqual(unique_paths(paths), [r"C:\Facturi\a.pdf", r"C:\Facturi\b.pdf"])


class TestOriginalNames(unittest.TestCase):
    def test_temp_names_are_mapped_back_in_one_pass(self):
        thread = ProcessingThread([], 0.6)
        thread.temp_names = {"0000_0001_X.PDF": "0001_X.PDF", "0001_X.PDF": "X.PDF"}
        self.assertEqual(thread._with_original_names("0000_0001_X.PDF apare și în 0001_X.PDF"),
                         "0001_X.PDF apare și în X.PDF")


if __name__ == "__main__":
    unittest.main()
