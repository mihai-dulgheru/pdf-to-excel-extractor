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
            os.rmdir(os.path.dirname(output_path))  # the staging folder; rmdir only ever removes an empty one
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


class TestOutputLocation(unittest.TestCase):
    def test_output_is_staged_in_a_folder_of_its_own_not_in_the_working_directory(self):
        """A fixed name in a shared folder could overwrite the workbook the user saved there earlier today."""
        import tempfile
        from unittest.mock import patch
        from modules.invoice_processor import InvoiceProcessor

        def one_row(path):
            return [{"company": "X", "invoice_number": "1", "nc8_code": "94019920", "origin": "RO",
                     "destination": "DE", "invoice_value_eur": 10, "net_weight": 1, "shipment_date": "01.01.2026",
                     "exchange_rate": 5.0, "value_ron": 50, "vat_number": "DE1", "delivery_location": 2093,
                     "delivery_condition": "FCA"}]

        with TemporaryDirectory() as root, TemporaryDirectory() as cwd:
            pdf = os.path.join(root, "x.pdf")
            with open(pdf, "wb") as f:
                f.write(b"not a pdf")
            previous = os.getcwd()
            os.chdir(cwd)
            try:
                with patch.object(InvoiceProcessor, "_process_single_invoice", staticmethod(one_row)):
                    output_path, _, _ = ProcessingThread([pdf], 0.6).process_invoices()
            finally:
                os.chdir(previous)
            staging = os.path.dirname(output_path)
            # clean up only what this test can prove the code created: a fresh folder directly under the temp dir
            in_temp = os.path.normcase(os.path.dirname(staging)) == os.path.normcase(tempfile.gettempdir())
            if in_temp:
                self.addCleanup(os.rmdir, staging)
                self.addCleanup(os.remove, output_path)
            self.assertTrue(in_temp, output_path)
            self.assertTrue(os.path.exists(output_path))
            self.assertEqual(os.path.basename(output_path), ProcessingThread([], 0.6).generate_excel_filename())
            self.assertNotEqual(os.path.normcase(staging), os.path.normcase(cwd))
            self.assertEqual(os.listdir(staging), [os.path.basename(output_path)])


class TestDiscardStaging(unittest.TestCase):
    def test_removes_only_a_staging_folder_the_app_created(self):
        from tempfile import mkdtemp
        from qt_ui import discard_staging
        staging = mkdtemp(prefix="pdf_to_excel_")
        with open(os.path.join(staging, "x.xlsx"), "wb") as f:
            f.write(b"1")
        discard_staging(os.path.join(staging, "x.xlsx"))
        self.assertFalse(os.path.exists(staging))
        with TemporaryDirectory() as other:  # any other folder: only the workbook goes, never the folder
            for name in ("x.xlsx", "keep.txt"):
                with open(os.path.join(other, name), "wb") as f:
                    f.write(b"1")
            discard_staging(os.path.join(other, "x.xlsx"))
            self.assertFalse(os.path.exists(os.path.join(other, "x.xlsx")))
            self.assertTrue(os.path.exists(os.path.join(other, "keep.txt")))


if __name__ == "__main__":
    unittest.main()
