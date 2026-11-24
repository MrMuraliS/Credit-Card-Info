import json
import subprocess
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from server import cards_from_workbook


class AddOnMappingTests(unittest.TestCase):
    def make_workbook(self, include_add_on: bool) -> Path:
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        path = Path(temporary.name) / "cards.xlsx"
        workbook = Workbook()
        summary = workbook.active
        summary.title = "Card Summary"
        summary.append(["Card ID", "Card Name", "Currency", "Standard Annual Fee"])
        summary.append(["TEST", "Test card", "INR", "INR 500"])
        rules = workbook.create_sheet("Benefits & Rules")
        rules.append([
            "Card ID", "Card Name", "Category", "Benefit / charge / rule",
            "Amount or rate", "Eligibility / spend criteria",
            "Exclusions / caps / limitations", "Verification status", "Source IDs",
        ])
        rules.append([
            "TEST", "Test card", "Movies", "Primary movie discount", "10%",
            "Primary condition", "Primary cap", "Verified", "TEST-S01",
        ])
        if include_add_on:
            rules.append([
                "TEST", "Test card", "Add-on", "Movies / BOGO and spend", "100%",
                "Add-on condition", "Cap sharing unverified", "Partial", "TEST-S01",
            ])
        sources = workbook.create_sheet("Source Register")
        sources.append([
            "Source ID", "Card ID(s)", "Official document/page title", "Publisher",
            "Official URL", "Document date / effective date", "Relevant section / page",
            "Facts checked", "Applicability / limitation", "Access date",
        ])
        sources.append(["TEST-S01", "TEST", "Terms", "Issuer", "https://issuer.example/terms"])
        workbook.save(path)
        workbook.close()
        return path

    def test_add_on_is_separate_and_does_not_affect_ranking(self):
        card = cards_from_workbook(self.make_workbook(True))[0]
        self.assertEqual(len(card["benefits"]), 1)
        self.assertEqual(len(card["addOn"]), 1)
        self.assertEqual(card["addOn"][0]["status"], "partial")
        self.assertEqual(card["addOn"][0]["sources"], "TEST-S01")
        self.assertIn("Cap sharing unverified", card["addOn"][0]["condition"])
        self.assertEqual(card["sources"][0][2], "https://issuer.example/terms")
        self.assertEqual(card["categoryRewardScores"]["movie"]["rate"], 10)

    def test_workbook_without_add_on_rows_returns_empty_list(self):
        card = cards_from_workbook(self.make_workbook(False))[0]
        self.assertEqual(card["addOn"], [])
        self.assertEqual(len(card["benefits"]), 1)

    @unittest.skipUnless(
        Path("/System/Library/Frameworks/JavaScriptCore.framework/Versions/A/Helpers/jsc").exists(),
        "JavaScriptCore command-line interpreter is unavailable",
    )
    def test_old_server_response_keeps_cards_visible_and_requests_restart(self):
        card = cards_from_workbook(self.make_workbook(True))[0]
        old_card = dict(card)
        old_card["benefits"] = card["benefits"] + card["addOn"]
        del old_card["addOn"]
        html = (Path(__file__).parent / "index.html").read_text()
        script = html.split("<script>", 1)[1].split("</script>", 1)[0]
        script = script.rsplit("    loadCards();", 1)[0]
        harness = """
        function assert(value, message) { if (!value) throw new Error(message); }
        const elements = new Map();
        function element(id) {
          if (!elements.has(id)) elements.set(id, {
            id, value: '', innerHTML: '', textContent: '', hidden: false,
            tagName: 'SELECT',
            classList: { contains() { return false; }, toggle() {}, remove() {} },
            addEventListener() {}, add() {}, reset() {}, scrollIntoView() {},
            querySelectorAll() { return []; }
          });
          return elements.get(id);
        }
        const document = { querySelector(selector) { return element(selector); } };
        const Option = function (text, value) { this.text = text; this.value = value; };
        let responseCards = OLD_CARDS;
        const fetch = async () => ({ ok: true, json: async () => ({ cards: responseCards }) });
        APP_SCRIPT
        loadCards().then(async () => {
          assert(cards.length === 1, 'Old API lost card');
          assert(element('#resultsList').innerHTML.includes('card-row'), 'Card not rendered');
          assert(element('#dataStatus').textContent.includes('restart server'), 'Missing warning');
          state.activeTab = 'addOn';
          renderDetail(cards[0]);
          assert(detailContent.innerHTML.includes('older API'), 'Missing detail restart guidance');
          state.compareIds = new Set([cards[0].id]);
          renderCompare();
          assert(element('#compareTable').innerHTML.includes('Restart the dashboard server'), 'Missing comparison guidance');
          responseCards = NEW_CARDS;
          await loadCards();
          assert(cards[0].addOn.length === 1, 'Current API add-on not loaded');
          assert(!element('#dataStatus').textContent.includes('restart'), 'Unexpected warning');
          state.activeTab = 'addOn';
          renderDetail(cards[0]);
          assert(detailContent.innerHTML.includes('Add-on condition'), 'Add-on detail not rendered');
          print('PASS');
        }).catch((error) => { print('FAIL: ' + error.message); });
        """.replace("OLD_CARDS", json.dumps([old_card])).replace(
            "NEW_CARDS", json.dumps([card])
        ).replace("APP_SCRIPT", script)
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "check.js"
            path.write_text(harness)
            result = subprocess.run(
                ["/System/Library/Frameworks/JavaScriptCore.framework/Versions/A/Helpers/jsc", str(path)],
                capture_output=True,
                text=True,
                check=True,
            )
        self.assertIn("PASS", result.stdout, result.stdout + result.stderr)
        self.assertNotIn("FAIL", result.stdout)


if __name__ == "__main__":
    unittest.main()
