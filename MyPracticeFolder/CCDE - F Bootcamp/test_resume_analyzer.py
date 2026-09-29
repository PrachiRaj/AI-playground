import csv
import tempfile
import tkinter as tk
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock

import resume_analyzer as ra

FAKE_KEY = "sk-ant-api03-FAKEkey_1234567890-abcdef"
HOSTILE = '<font color="red">Bob</font> & <b>Co</b> <para> ]]> <img src="x"/> "quotes" \'apos\''


def make_result(**overrides):
    result = {
        "file": "cv <1>.pdf", "candidate_name": HOSTILE, "match_percentage": 90, "status": "Selected", "error": "",
        "summary": HOSTILE, "experience_fit": "5 < 7 years & rising",
        "matched_skills": [HOSTILE, "C++ & <STL>"], "missing_skills": ["</b>"], "strengths": ["a > b"], "gaps": [],
    }
    result.update(overrides)
    return result


class ScreenerTestCase(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()
        with mock.patch.object(ra, "load_env_key", return_value=(None, None, None)):
            self.app = ra.ResumeScreenerApp(self.root)
        self.tmp = Path(tempfile.mkdtemp())
        self.patches = [
            mock.patch.object(ra.messagebox, "showinfo"),
            mock.patch.object(ra.messagebox, "showerror"),
        ]
        self.showinfo, self.showerror = (p.start() for p in self.patches)

    def tearDown(self):
        for p in self.patches:
            p.stop()
        self.root.destroy()

    def load_results(self, results, jd="Need <Python> & SQL"):
        self.app.results = results
        self.app.run_info = {"threshold": 85, "date": datetime.now(), "jd": jd}
        self.app.build_results_tab()

    def export(self, method, ext):
        out = self.tmp / f"report{ext}"
        with mock.patch.object(ra.filedialog, "asksaveasfilename", return_value=str(out)):
            method()
        return out


class PdfEscapingTests(ScreenerTestCase):
    def test_markup_in_ai_fields_does_not_break_pdf(self):
        self.load_results([
            make_result(),
            make_result(status="Rejected", match_percentage=40, candidate_name="Amy & <Lee>"),
            make_result(status="Error", match_percentage=0, error=HOSTILE),
        ], jd=HOSTILE)
        out = self.export(self.app.export_pdf, ".pdf")
        self.showerror.assert_not_called()
        self.assertTrue(out.read_bytes().startswith(b"%PDF"))

    def test_non_latin_names_do_not_break_pdf(self):
        self.load_results([make_result(candidate_name="José Müller 李雷", matched_skills=["Café ☕"])])
        self.export(self.app.export_pdf, ".pdf")
        self.showerror.assert_not_called()


class KeyRedactionTests(ScreenerTestCase):
    def test_redact_removes_key_patterns_and_exact_key(self):
        self.assertNotIn(FAKE_KEY, ra.redact(f"bad header x-api-key: {FAKE_KEY}"))
        self.assertNotIn("custom-secret", ra.redact("value custom-secret here", key="custom-secret"))
        self.assertEqual(ra.redact("nothing secret"), "nothing secret")

    def test_key_in_error_never_reaches_ui_or_exports(self):
        finished = {}
        self.app.ui = lambda fn, *a, **k: finished.update(args=a) if fn == self.app.finish else None
        resume = self.tmp / "cv.txt"
        resume.write_text("Jane Doe, Python developer")
        with mock.patch.object(ra, "analyze_resume", side_effect=RuntimeError(f"upstream echoed {FAKE_KEY}")):
            self.app.run_screening(FAKE_KEY, "JD text", [str(resume)], 85)

        results = finished["args"][0]
        self.assertEqual(results[0]["status"], "Error")
        self.assertNotIn(FAKE_KEY, results[0]["error"])

        self.load_results(results)
        csv_text = self.export(self.app.export_csv, ".csv").read_text(encoding="utf-8-sig")
        self.assertNotIn(FAKE_KEY, csv_text)
        self.assertIn("[REDACTED]", next(csv.reader([csv_text.splitlines()[1]]))[-1])

    def test_fatal_error_is_redacted(self):
        finished = {}
        self.app.ui = lambda fn, *a, **k: finished.update(args=a) if fn == self.app.finish else None
        resume = self.tmp / "cv.txt"
        resume.write_text("Jane Doe")
        with mock.patch.object(ra, "analyze_resume", side_effect=ra.FatalAPIError(f"denied for {FAKE_KEY}")):
            self.app.run_screening(FAKE_KEY, "JD text", [str(resume)], 85)
        self.assertNotIn(FAKE_KEY, finished["args"][1])


class EnvParsingTests(unittest.TestCase):
    def test_formats(self):
        cases = {
            "ANTHROPIC_API_KEY=sk-ant-a": "sk-ant-a",
            "ANTHROPIC_API_KEY:sk-ant-b": "sk-ant-b",
            'export CLAUDE_API_KEY="sk-ant-c"': "sk-ant-c",
            "# comment\nsk-ant-d": "sk-ant-d",
            "OTHER=1": None,
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                self.assertEqual(ra.parse_env_key(text), expected)


if __name__ == "__main__":
    unittest.main()
