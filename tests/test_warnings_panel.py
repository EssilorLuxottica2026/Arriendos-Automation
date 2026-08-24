from pathlib import Path
import unittest

from flask import Flask, render_template


class WarningsPanelTests(unittest.TestCase):
    def test_panel_renders_compact_summary_and_all_warning_messages(self):
        project_root = Path(__file__).resolve().parents[1]
        app = Flask(__name__, template_folder=str(project_root / "templates"))
        warnings = [
            "CECO 6626: se ignoro un contrato con estado REEMPLAZO.",
            "Se omitieron 2 facturas porque no tienen XML coincidente.",
        ]

        with app.test_request_context("/"):
            html = render_template("_warnings_panel.html", process_warnings=warnings)

        self.assertIn("2 avisos requieren", html)
        self.assertIn("Revisar avisos", html)
        self.assertIn("processWarningsDrawer", html)
        self.assertIn(warnings[0], html)
        self.assertIn(warnings[1], html)

    def test_panel_renders_nothing_without_warnings(self):
        project_root = Path(__file__).resolve().parents[1]
        app = Flask(__name__, template_folder=str(project_root / "templates"))

        with app.test_request_context("/"):
            html = render_template("_warnings_panel.html", process_warnings=[])

        self.assertEqual(html.strip(), "")


if __name__ == "__main__":
    unittest.main()
