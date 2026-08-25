import email
import importlib.util
import sys
import tempfile
import unittest
from datetime import date
from email.message import EmailMessage
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))
import download_and_register as core


class FakePdfPage:
    def extract_text(self):
        return """增值税电子普通发票
发票号码：12345678901234567890
开票日期：2026年08月24日
购买方 名称：示例购买方 统一社会信用代码：X
销售方 名称：示例销售方 统一社会信用代码：Y
*信息技术服务* 100.00 6.00
合 计 ¥100.00 ¥6.00
价 税 合 计（小写）¥106.00
"""


class FakePdf:
    pages = [FakePdfPage()]
    def __enter__(self): return self
    def __exit__(self, *_): return False


class CoreTests(unittest.TestCase):
    def test_validate_config(self):
        cfg = core.validate_config({"IMAP_HOST": "imap.example.com", "IMAP_USERNAME": "u@example.com", "IMAP_PASSWORD": "secret"})
        self.assertEqual(cfg["IMAP_PORT"], "993")
        self.assertEqual(cfg["IMAP_MAILBOX"], "INBOX")
        with self.assertRaises(ValueError):
            core.validate_config({"IMAP_HOST": "bad host", "IMAP_USERNAME": "u", "IMAP_PASSWORD": "x"})

    def test_candidate_filter(self):
        msg = EmailMessage()
        msg["Subject"] = "您的电子发票"
        msg["From"] = "billing@example.com"
        msg["Date"] = "Mon, 24 Aug 2026 10:00:00 +0800"
        msg.set_content("请查收")
        msg.add_attachment(b"pdf-bytes", maintype="application", subtype="pdf", filename="invoice.pdf")
        msg.add_attachment(b"other", maintype="application", subtype="octet-stream", filename="manual.docx")
        items = list(core.iter_candidates(msg, "1", core.DEFAULT_KEYWORDS, False, core.DOWNLOADABLE_EXTENSIONS))
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0][0].filename, "invoice.pdf")

    def test_parse_invoice_and_summary(self):
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "invoice.pdf"
            p.write_bytes(b"fake")
            with patch.object(core.pdfplumber, "open", return_value=FakePdf()):
                invoice = core.parse_invoice_pdf(p)
            self.assertEqual(invoice["发票号码"], "12345678901234567890")
            self.assertEqual(invoice["价税合计"], 106.0)
            self.assertEqual(invoice["校验结果"], "通过")
            r1 = core.MailAttachment("1", "a", "s", "d", "a.pdf", ".pdf", invoice=invoice, status="已解析")
            r2 = core.MailAttachment("2", "a", "s", "d", "b.pdf", ".pdf", invoice=dict(invoice), status="已解析")
            summary = core.build_summary([r1, r2])
            self.assertEqual(summary["duplicate_invoice_numbers"], 1)
            self.assertNotIn("password", str(summary).lower())

    def test_local_folder_mode(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            source, output = root / "source", root / "output"
            source.mkdir()
            (source / "one.pdf").write_bytes(b"fake")
            with patch.object(core, "parse_invoice_pdf", return_value={"发票号码": "1", "开票日期": "", "销方名称": "A", "购方名称": "B", "价税合计": 10.0, "合计金额": 9.0, "合计税额": 1.0, "项目明细": [], "解析状态": "已解析", "校验结果": "通过"}):
                records, summary = core.run_local_folder_job(source, output, True)
            self.assertEqual(len(records), 1)
            self.assertEqual(summary["source_mode"], "local_folder")
            self.assertTrue(Path(summary["output_file"]).is_file())
            self.assertTrue((source / "one.pdf").is_file())

    def test_workbook_marks_duplicate(self):
        with tempfile.TemporaryDirectory() as td:
            invoice = {"发票号码": "10001", "开票日期": "2026-08-24", "销方名称": "A", "购方名称": "B", "价税合计": 10.0, "合计金额": 9.0, "合计税额": 1.0, "项目明细": [], "解析状态": "已解析", "校验结果": "通过"}
            records = [core.MailAttachment(str(i), "s", "t", "d", f"{i}.pdf", ".pdf", invoice=dict(invoice), status="已解析") for i in (1, 2)]
            out = Path(td) / "out.xlsx"
            core.create_workbook(records, out)
            import openpyxl
            wb = openpyxl.load_workbook(out)
            self.assertIn("重复发票号码", wb["发票汇总"].cell(2, 14).value)


if __name__ == "__main__":
    unittest.main()
