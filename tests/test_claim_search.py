import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from CC.claim_search import ClaimSearcher


class ClaimSearcherTest(unittest.TestCase):
    """Tests for ``ClaimSearcher``."""

    def setUp(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.append(["Complaint", "ID"])
        ws.append(["Engine noise at startup", 1])
        ws.append(["Brake squeal when hot", 2])
        ws.append(["Engine stalls at idle", 3])
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        wb.save(tmp.name)
        tmp.close()
        self.temp_path = Path(tmp.name)
        self.searcher = ClaimSearcher(str(self.temp_path))

    def tearDown(self) -> None:
        self.temp_path.unlink()

    def test_find_similar_finds_match(self) -> None:
        results = self.searcher.find_similar(
            "engine noise when starting",
            threshold=0.7,
        )
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["Complaint"], "Engine noise at startup")
        self.assertEqual(results[0]["ID"], 1)

    def test_find_similar_respects_threshold(self) -> None:
        results = self.searcher.find_similar("engine noise when starting", threshold=0.9)
        self.assertEqual(results, [])


if __name__ == "__main__":
    unittest.main()
