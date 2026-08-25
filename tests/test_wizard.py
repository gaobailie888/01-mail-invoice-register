import http.client
import importlib.util
import json
import sys
import threading
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))
spec = importlib.util.spec_from_file_location("setup_wizard", SCRIPTS / "setup_wizard.py")
wizard = importlib.util.module_from_spec(spec)
spec.loader.exec_module(wizard)


class WizardTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = wizard.ThreadingHTTPServer((wizard.HOST, 0), wizard.Handler)
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.port = cls.server.server_port

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join(timeout=3)

    def request(self, method, path, body=None, headers=None):
        conn = http.client.HTTPConnection("127.0.0.1", self.port, timeout=5)
        conn.request(method, path, body=body, headers=headers or {})
        res = conn.getresponse()
        data = res.read()
        conn.close()
        return res.status, data

    def test_home_does_not_contain_password(self):
        status, body = self.request("GET", "/", headers={"Host": f"127.0.0.1:{self.port}"})
        self.assertEqual(status, 200)
        text = body.decode("utf-8")
        self.assertNotIn("replace-with-an-app-password", text)
        self.assertIn(wizard.TOKEN, text)

    def test_rejects_foreign_host(self):
        status, _ = self.request("GET", "/", headers={"Host": "evil.example"})
        self.assertEqual(status, 403)

    def test_home_contains_buyer_options_without_private_rules(self):
        status, body = self.request("GET", "/", headers={"Host": f"127.0.0.1:{self.port}"})
        text = body.decode("utf-8")
        self.assertEqual(status, 200)
        self.assertIn("organize_files", text)
        self.assertIn("buyer_sheets", text)
        self.assertNotIn("真实客户公司", text)
        self.assertNotIn("真实业务文件", text)

    def test_rejects_missing_csrf_token(self):
        payload = json.dumps({"host": "imap.example.com", "username": "u@example.com", "password": "top-secret"})
        status, body = self.request("POST", "/api/test", payload, {"Host": f"127.0.0.1:{self.port}", "Content-Type": "application/json"})
        self.assertEqual(status, 403)
        self.assertNotIn(b"top-secret", body)


if __name__ == "__main__":
    unittest.main()
