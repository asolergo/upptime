"""Pruebas de exportación de ejecuciones de agente sin inventario."""

import contextlib
import csv
import importlib.util
import io
from pathlib import Path
import sqlite3
import tempfile
import unittest


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "consolidar_inventario.py"
SPEC = importlib.util.spec_from_file_location("consolidar_inventario_agent_export", SCRIPT_PATH)
inventory = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(inventory)


class AgentExportTests(unittest.TestCase):
    def test_exports_only_unmatched_latest_finding_with_detail(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "remediacion_agentica.csv"
            output = root / "Inventario.xlsx"
            rows = [
                {
                    "id": "exec-1", "agent_name": "agent-a", "user_id": "user-a",
                    "action": "scan", "timestamp": "2026-10-28T10:00:00Z",
                    "status": "completed", "duration_ms": "100", "repository": "repo-a",
                    "details": '{"framework":"cloud","project":"project-a","findings":'
                               '[{"finding_id":"10","resolution":"fixed","severity":"high",'
                               '"product":"product-a","summary":"matched"},'
                               '{"finding_id":"20","resolution":"old","summary":"old"}]}',
                },
                {
                    "id": "exec-2", "agent_name": "agent-b", "user_id": "user-b",
                    "action": "remediate", "timestamp": "2026-10-29T10:00:00Z",
                    "status": "completed", "duration_ms": "200", "repository": "repo-b",
                    "details": '{"framework":"batch","findings":'
                               '[{"finding_id":"20","resolution":"fixed","severity":"medium",'
                               '"product":"product-b","summary":"latest"}]}',
                },
            ]
            with source.open("w", encoding="utf-8", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
                writer.writeheader()
                writer.writerows(rows)

            connection = sqlite3.connect(":memory:")
            self.addCleanup(connection.close)
            inventory.create_agent_table(connection)
            with contextlib.redirect_stdout(io.StringIO()):
                inventory.process_agent_execution_file(connection, source)
            connection.execute('CREATE TABLE inventario ("id" TEXT)')
            connection.execute('INSERT INTO inventario VALUES ("10")')

            detail_path, count = inventory.write_unmatched_agent_findings_csv(connection, output)
            index = connection.execute(
                "SELECT name FROM sqlite_master WHERE type='index' AND name=?",
                ("idx_inventario_id_text",),
            ).fetchone()
            latest_path = root / "logs" / "ejecuciones_agente_sin_inventario_ULTIMO.csv"
            self.assertEqual(count, 1)
            self.assertIsNotNone(index)
            self.assertTrue(detail_path.is_file())
            self.assertTrue(latest_path.is_file())
            with detail_path.open(encoding="utf-8-sig", newline="") as handle:
                reader = csv.DictReader(handle, delimiter=";")
                exported = list(reader)
            self.assertEqual(len(exported), 1)
            row = exported[0]
            self.assertEqual(row["finding_id"], "20")
            self.assertEqual(row["resolution"], "fixed")
            self.assertEqual(row["execution_id"], "exec-2")
            self.assertEqual(row["agent_name"], "agent-b")
            self.assertEqual(row["repository"], "repo-b")
            self.assertEqual(row["framework"], "batch")
            self.assertEqual(row["summary"], "latest")
            self.assertEqual(inventory.json.loads(row["detalle_finding_json"])["finding_id"], "20")


if __name__ == "__main__":
    unittest.main()
