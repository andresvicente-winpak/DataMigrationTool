"""Dependency-free structural checks for the MCO importer.

This test intentionally uses only the standard library so syntax regressions can
be diagnosed even on a workstation where the application's Python dependencies
have not loaded yet.
"""

import ast
from pathlib import Path
import unittest


IMPORTER_PATH = Path(__file__).parents[1] / "modules" / "mco_importer.py"


class TestMCOImporterSource(unittest.TestCase):
    def test_importer_parses_and_has_one_classification_path(self):
        source = IMPORTER_PATH.read_text(encoding="utf-8")
        tree = ast.parse(source, filename=str(IMPORTER_PATH))

        importer = next(
            node
            for node in tree.body
            if isinstance(node, ast.ClassDef) and node.name == "MCOImporter"
        )
        classifiers = [
            node
            for node in importer.body
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
            and node.name == "_classify_rule"
        ]

        self.assertEqual(1, len(classifiers))


if __name__ == "__main__":
    unittest.main()
