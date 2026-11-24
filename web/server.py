from __future__ import annotations

import argparse
import json
import re
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
WEB_DIR = Path(__file__).resolve().parent
WORKBOOK = ROOT / "output" / "credit-card-comparison.xlsx"


def _records(sheet: Any) -> list[dict[str, Any]]:
    rows = sheet.iter_rows(values_only=True)
    headers = next(rows, ())
    return [
        dict(zip(headers, row))
        for row in rows
        if any(value is not None for value in row)
    ]


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _money_number(value: str) -> float | None:
    no_fee = r"\b(?:nil|no annual fee|lifetime free)\b"
    if re.search(no_fee, value, re.IGNORECASE):
        return 0.0
    pattern = r"(?:INR|Rs\.?|₹)?\s*([0-9][0-9,]*(?:\.\d+)?)"
    match = re.search(pattern, value, re.IGNORECASE)
    return float(match.group(1).replace(",", "")) if match else None


def _money_label(value: str, currency: str) -> str:
    amount = _money_number(value)
    if amount is None:
        return "Unverified"
    formatted = f"{amount:,.0f}" if amount.is_integer() else f"{amount:,.2f}"
    return f"{currency} {formatted}"


def _compact_rate(value: str) -> str:
    range_match = re.search(
        r"(\d+(?:\.\d+)?\s*%)\s*(?:-|–|to)\s*(\d+(?:\.\d+)?\s*%)",
        value,
        re.IGNORECASE,
    )
    if range_match:
        lower = range_match.group(1).replace(" ", "")
        upper = range_match.group(2).replace(" ", "")
        return f"{lower}–{upper} p.a."
    match = re.search(r"\d+(?:\.\d+)?\s*%", value)
    if not match:
        return value or "Unverified"
    return f"{match.group(0).replace(' ', '')} p.a."


def _compact_percent(value: str) -> str:
    match = re.search(r"\d+(?:\.\d+)?\s*%", value)
    if match:
        return match.group(0).replace(" ", "")
    if value.lower().startswith("unverified"):
        return "Unverified"
    return value or "Unverified"


def _compact_rewards(value: str, limit: int = 52) -> str:
    if not value:
        return "Unverified"
    if value.lower().startswith("unverified"):
        return "Everyday earn rate unverified"
    first_clause = value.split(";")[0].strip()
    if len(first_clause) <= limit:
        return first_clause
    return first_clause[: limit - 1].rsplit(" ", 1)[0] + "…"


def _sentences(value: str) -> list[str]:
    pattern = r"\n+|;\s*|(?<=[.!?])\s+(?=[A-Z₹])"
    parts = re.split(pattern, value)
    return [part.strip(" .;\n") for part in parts if part.strip(" .;\n")]


def _benefit_tags(rules: list[dict[str, Any]]) -> list[str]:
    category_labels = {
        "rewards": "Rewards",
        "milestone": "Rewards",
        "quarterly milestone": "Rewards",
        "cashback": "Cashback",
        "cash back": "Cashback",
        "partner cashback": "Cashback",
        "travel": "Travel",
        "travel benefits": "Travel",
        "airport lounge": "Lounges",
        "airport lounges": "Lounges",
        "airport lounge access": "Lounges",
        "lounge access": "Lounges",
        "domestic airport lounge": "Lounges",
        "fuel": "Fuel",
        "movies": "Movies",
        "movie": "Movies",
        "entertainment": "Entertainment",
        "monthly entertainment": "Entertainment",
        "dining": "Dining",
        "insurance": "Insurance",
        "fee waiver": "Fee waiver",
        "annual fee waiver": "Fee waiver",
        "upi": "UPI",
        "upi rewards": "UPI",
        "golf": "Golf",
        "shopping": "Shopping",
    }
    uncertain_status_terms = (
        "unverified", "conflict", "not verified", "partial", "ambiguous",
        "ambiguity", "discontinued", "discontinuation", "no longer",
    )
    tags = set()
    for rule in rules:
        status = _text(rule.get("Verification status")).lower()
        if not status.startswith("verified") or any(
            term in status for term in uncertain_status_terms
        ):
            continue
        category = re.sub(r"\s+", " ", _text(rule.get("Category")).lower())
        label = category_labels.get(category)
        if label:
            tags.add(label)

    order = (
        "Rewards", "Cashback", "Movies", "Entertainment", "Dining", "Travel",
        "Lounges", "Fuel", "UPI", "Shopping", "Golf", "Insurance",
        "Fee waiver",
    )
    return [label for label in order if label in tags][:4] or [
        "Terms unverified" if rules else "Credit card"
    ]


def _cashback_summary(
    rules: list[dict[str, Any]],
) -> tuple[float | None, str, str]:
    uncertain_status_terms = (
        "unverified", "conflict", "not verified", "partial", "ambiguous",
        "ambiguity", "discontinued", "discontinuation", "no longer",
    )
    candidates: list[tuple[float, str, str]] = []
    for rule in rules:
        status = _text(rule.get("Verification status")).lower()
        if not status.startswith("verified") or any(
            term in status for term in uncertain_status_terms
        ):
            continue
        category = _text(rule.get("Category")).lower()
        title = _text(rule.get("Benefit / charge / rule"))
        if category not in {"cashback", "partner cashback"}:
            continue
        amount = _text(rule.get("Amount or rate"))
        rates = [
            float(value)
            for value in re.findall(r"(?<!\d)(\d+(?:\.\d+)?)\s*%", amount)
        ]
        if rates:
            conditions = _text(rule.get("Eligibility / spend criteria"))
            candidates.append((max(rates), title, conditions))
    if not candidates:
        return None, "No verified percentage cashback rate recorded", ""
    rate, title, conditions = max(
        candidates, key=lambda candidate: candidate[0]
    )
    return rate, title, conditions


def _category_reward_scores(
    rules: list[dict[str, Any]],
) -> dict[str, dict[str, Any]]:
    focus_categories = {
        "movie": {
            "movies", "movie", "entertainment", "monthly entertainment"
        },
        "entertainment": {
            "movies", "movie", "entertainment", "monthly entertainment"
        },
        "cashback": {"cashback", "partner cashback"},
        "dining": {"dining"},
        "fuel": {"fuel"},
        "upi": {"upi", "upi rewards"},
        "lounge": {
            "airport lounge", "airport lounges", "airport lounge access",
            "lounge access",
        },
        "travel": {"travel", "travel benefits", "travel milestone"},
        "golf": {"golf"},
        "insurance": {"insurance"},
        "rewards": {
            "rewards", "milestone", "quarterly milestone", "welcome benefit"
        },
        "fee": {"fee waiver", "annual fee waiver"},
    }
    uncertain_status_terms = (
        "unverified", "conflict", "not verified", "partial", "ambiguous",
        "ambiguity", "discontinued", "discontinuation", "no longer",
    )
    fuel_point_redemption_value: float | None = None
    for rule in rules:
        category = _text(rule.get("Category")).lower()
        status = _text(rule.get("Verification status")).lower()
        amount = _text(rule.get("Amount or rate"))
        if category != "fuel" or not status.startswith("verified"):
            continue
        match = re.search(
            r"1\s+Fuel Point\s*=\s*(?:INR|Rs\.?|₹)\s*([0-9.]+)",
            amount,
            re.IGNORECASE,
        )
        if match:
            fuel_point_redemption_value = float(match.group(1))
            break
    scores: dict[str, dict[str, Any]] = {}
    for focus, categories in focus_categories.items():
        candidates: list[dict[str, Any]] = []
        for rule in rules:
            status = _text(rule.get("Verification status")).lower()
            uncertain_status = any(
                term in status for term in uncertain_status_terms
            )
            scope_limited_movie = (
                focus == "movie" and "scope ambiguity" in status
            )
            if not status.startswith("verified") or (
                uncertain_status and not scope_limited_movie
            ):
                continue
            category = re.sub(
                r"\s+", " ", _text(rule.get("Category")).lower()
            )
            title = _text(rule.get("Benefit / charge / rule"))
            amount = _text(rule.get("Amount or rate"))
            criteria = _text(rule.get("Eligibility / spend criteria"))
            limitations = _text(rule.get("Exclusions / caps / limitations"))
            if focus == "fuel" and "convenience fee" in title.lower():
                continue
            category_match = category in categories
            if focus in {"movie", "entertainment"}:
                category_match = category_match or any(
                    term in title.lower()
                    for term in ("movie", "bookmyshow", "district")
                )
            elif focus == "upi":
                category_match = category_match or "upi" in (
                    title + " " + amount
                ).lower()
            elif focus == "dining":
                category_match = category_match or any(
                    term in (title + " " + amount).lower()
                    for term in (
                        "dining", "restaurant", "food", "swiggy", "zomato"
                    )
                )
            elif focus == "travel":
                category_match = category_match or any(
                    term in title.lower()
                    for term in ("travel", "airline", "miles", "flight")
                )
            elif focus == "rewards":
                category_match = category_match or any(
                    term in title.lower()
                    for term in ("reward", "point", "mile")
                )
            if not category_match:
                continue

            rates = [
                float(value)
                for value in re.findall(r"(?<!\d)(\d+(?:\.\d+)?)\s*%", amount)
            ]
            if focus == "fuel" and not rates:
                fuel_points = re.search(
                    r"(\d+)\s+Fuel Points?\s+per\s+(?:INR|Rs\.?|₹)\s*100",
                    amount,
                    re.IGNORECASE,
                )
                point_value_match = re.search(
                    r"1\s+Fuel Point\s*=\s*(?:INR|Rs\.?|₹)\s*([0-9.]+)",
                    amount,
                    re.IGNORECASE,
                )
                point_value = (
                    float(point_value_match.group(1))
                    if point_value_match
                    else fuel_point_redemption_value
                )
                if fuel_points and point_value is not None:
                    redemption_value = (
                        float(fuel_points.group(1))
                        * point_value
                    )
                    monthly_points = re.search(
                        r"max(?:imum)?\s+([\d,]+)\s+Fuel Points?\s+per month",
                        amount,
                        re.IGNORECASE,
                    )
                    monthly_value = (
                        float(monthly_points.group(1).replace(",", ""))
                        * point_value
                        if monthly_points
                        else 0.0
                    )
                    candidates.append({
                        "rankGroup": 4,
                        "rate": redemption_value,
                        "cap": monthly_value,
                        "title": title,
                        "conditions": criteria or limitations,
                        "unit": "redemption",
                    })
                    continue
            if focus == "upi":
                upi_rates = [
                    float(value)
                    for value in re.findall(
                        (
                            r"(\d+(?:\.\d+)?)\s*%\s*"
                            r"(?:cashback\s*)?(?:on\s*)?(?:UPI|scan\s*&\s*pay)"
                        ),
                        amount,
                        re.IGNORECASE,
                    )
                ]
                upi_rates.extend(
                    float(value)
                    for value in re.findall(
                        (
                            r"(?:UPI|scan\s*&\s*pay)[^%;.]{0,50}?"
                            r"(\d+(?:\.\d+)?)\s*%"
                        ),
                        amount,
                        re.IGNORECASE,
                    )
                )
                if upi_rates:
                    rates = upi_rates
            if focus in {"movie", "entertainment"} and re.search(
                r"\b(?:bogo|buy one get one|second ticket|free tickets?)\b",
                " ".join((title, amount)).lower(),
            ):
                rates.append(100.0)
            if not rates and focus in {"movie", "entertainment"} and re.search(
                r"\b(?:discount|cashback|value.back)\b",
                " ".join((title, amount)).lower(),
            ):
                caps = [
                    float(value.replace(",", ""))
                    for value in re.findall(
                        r"(?:INR|Rs\.?|₹)\s*([0-9][0-9,]*(?:\.\d+)?)", amount,
                        re.IGNORECASE,
                    )
                ]
                thresholds = [
                    float(value.replace(",", ""))
                    for value in re.findall(
                        (
                            r"(?:minimum|at least)\s+(?:eligible\s+)?"
                            r"(?:spend|purchase|transaction|order)?\s*"
                            r"(?:of\s*)?(?:INR|Rs\.?|₹)\s*"
                            r"([0-9][0-9,]*(?:\.\d+)?)"
                        ),
                        criteria,
                        re.IGNORECASE,
                    )
                ]
                if caps and thresholds:
                    rates.append(min(caps) / min(thresholds) * 100)
            caps = [
                float(value.replace(",", ""))
                for value in re.findall(
                    r"(?:INR|Rs\.?|₹)\s*([0-9][0-9,]*(?:\.\d+)?)", amount,
                    re.IGNORECASE,
                )
            ]
            cap = max(caps, default=0.0)
            amount_text = " ".join((amount, limitations)).lower()
            if re.search(
                r"\btwice\s+(?:a\s+)?(?:month|monthly)\b", amount_text
            ):
                cap *= 2
            else:
                ticket_count = re.search(
                    (
                        r"(?:up to\s*)?(\d+)\s+(?:free\s+)?tickets?"
                        r"\s*(?:/|per|every)?\s*(?:calendar\s*)?month"
                    ),
                    amount_text,
                )
                if ticket_count and "per ticket" in amount_text:
                    cap *= float(ticket_count.group(1))
            conditions = criteria or limitations
            if rates:
                candidates.append({
                    "rankGroup": 3,
                    "rate": max(rates),
                    "cap": cap,
                    "title": title,
                    "conditions": conditions,
                    "unit": "percent",
                })
                continue

            if focus == "lounge":
                visits = [
                    float(value)
                    for value in re.findall(
                        (
                            r"(\d+)\s+(?:complimentary\s+)?"
                            r"(?:visits?|vouchers?|passes?|accesses?)"
                        ),
                        amount,
                        re.IGNORECASE,
                    )
                ]
                visits.extend(
                    float(value)
                    for value in re.findall(
                        (
                            r"(\d+)\s+.{0,35}?"
                            r"(?:visits?|vouchers?|passes?|accesses?)"
                        ),
                        amount,
                        re.IGNORECASE,
                    )
                )
                visits.extend(
                    float(value)
                    for value in re.findall(
                        r"(\d+)\s+per\s+(?:calendar\s+)?quarter",
                        amount,
                        re.IGNORECASE,
                    )
                )
                if visits:
                    candidates.append({
                        "rankGroup": 2, "rate": max(visits), "cap": 0,
                        "title": title, "conditions": conditions,
                        "unit": "visits",
                    })
                    continue

            if focus == "golf":
                rounds = [
                    float(value)
                    for value in re.findall(
                        (
                            r"(?:up to\s*)?(\d+)\s+"
                            r"(?:golf\s+)?(?:rounds?|lessons?)"
                        ),
                        amount,
                        re.IGNORECASE,
                    )
                ]
                rounds.extend(
                    float(value)
                    for value in re.findall(
                        r"(?:up to\s*)?(\d+)\s+per\s+(?:calendar\s+)?month",
                        amount,
                        re.IGNORECASE,
                    )
                )
                if rounds:
                    candidates.append({
                        "rankGroup": 2, "rate": max(rounds), "cap": 0,
                        "title": title, "conditions": conditions,
                        "unit": "rounds",
                    })
                    continue

            if focus == "insurance":
                crores = [
                    float(value)
                    for value in re.findall(
                        r"(?:INR|Rs\.?|₹)?\s*(\d+(?:\.\d+)?)\s*crore",
                        amount,
                        re.IGNORECASE,
                    )
                ]
                amounts = [
                    float(value.replace(",", ""))
                    for value in re.findall(
                        r"(?:INR|Rs\.?|₹)\s*([0-9][0-9,]*(?:\.\d+)?)",
                        amount,
                        re.IGNORECASE,
                    )
                ]
                lakhs = [
                    float(value.replace(",", "")) * 100_000
                    for value in re.findall(
                        r"(?:INR|Rs\.?|₹)\s*([0-9][0-9,]*(?:\.\d+)?)\s*lakh",
                        amount,
                        re.IGNORECASE,
                    )
                ]
                coverage = max(crores, default=0) * 10_000_000
                coverage = max(
                    coverage, max(lakhs, default=0), max(amounts, default=0)
                )
                if coverage:
                    candidates.append({
                        "rankGroup": 2, "rate": coverage, "cap": 0,
                        "title": title, "conditions": conditions,
                        "unit": "coverage",
                    })
                    continue

            if focus in {"rewards", "travel", "dining"}:
                points = re.search(
                    r"(\d+(?:\.\d+)?)\s*(?:EDGE\s*)?(?:Reward\s*)?"
                    r"(?:RPs?|Points?|Miles?)\s*(?:per|/)\s*"
                    r"(?:INR|Rs\.?|₹)\s*([\d,]+)",
                    amount,
                    re.IGNORECASE,
                )
                if points:
                    points_per_100 = float(points.group(1)) * 100 / float(
                        points.group(2).replace(",", "")
                    )
                    candidates.append({
                        "rankGroup": 1, "rate": points_per_100, "cap": 0,
                        "title": title, "conditions": conditions,
                        "unit": "points",
                    })
                    continue

        if candidates:
            scores[focus] = max(
                candidates,
                key=lambda item: (
                    item["rankGroup"], item["rate"], item["cap"]
                ),
            )
    return scores


def cards_from_workbook(path: Path = WORKBOOK) -> list[dict[str, Any]]:
    if not path.exists():
        raise FileNotFoundError(f"Comparison workbook not found: {path}")
    workbook = load_workbook(path, data_only=True, read_only=True)
    required = {"Card Summary", "Benefits & Rules", "Source Register"}
    if not required.issubset(workbook.sheetnames):
        sheet_names = ", ".join(sorted(required))
        raise ValueError(f"Workbook must contain sheets: {sheet_names}")

    summaries = _records(workbook["Card Summary"])
    rules_by_card: dict[str, list[dict[str, Any]]] = {}
    for rule in _records(workbook["Benefits & Rules"]):
        rules_by_card.setdefault(_text(rule.get("Card ID")), []).append(rule)
    sources_by_card: dict[str, list[dict[str, Any]]] = {}
    for source in _records(workbook["Source Register"]):
        for card_id in _text(source.get("Card ID(s)")).split(";"):
            card_id = card_id.strip()
            if card_id:
                sources_by_card.setdefault(card_id, []).append(source)

    cards = []
    for summary in summaries:
        card_id = _text(summary.get("Card ID"))
        if not card_id:
            continue
        card_rules = rules_by_card.get(card_id, [])
        card_sources = sources_by_card.get(card_id, [])
        fees = []
        benefits = []
        add_on = []
        eligibility = []
        for rule in card_rules:
            category = _text(rule.get("Category"))
            title = (
                _text(rule.get("Benefit / charge / rule"))
                or category
                or "Card term"
            )
            amount = _text(rule.get("Amount or rate"))
            criteria = _text(rule.get("Eligibility / spend criteria"))
            limitations = _text(rule.get("Exclusions / caps / limitations"))
            status_source = _text(rule.get("Verification status"))
            source_ids = _text(rule.get("Source IDs"))
            detail_parts = [criteria, limitations]
            detail = " ".join(part for part in detail_parts if part)
            low_category = category.lower()
            low_title = title.lower()
            fee_categories = {"fees", "interest", "other fees"}
            fee_titles = {
                "joining fee",
                "annual fee",
                "purchase/revolving apr",
                "standard joining fee",
                "standard annual fee",
            }
            if low_category in fee_categories or low_title in fee_titles:
                fee_parts = (amount, criteria, limitations)
                fee_value = " ".join(part for part in fee_parts if part)
                fees.append([title, fee_value or "See official source."])
                continue
            if "eligibility" in low_category:
                eligibility_parts = (amount, criteria, limitations)
                eligibility_text = " ".join(
                    part for part in eligibility_parts if part
                )
                eligibility.extend(
                    _sentences(eligibility_text)
                )
                continue
            status_low = status_source.lower()
            unverified_terms = (
                "unverified", "conflict", "unavailable", "unclear"
            )
            partial_terms = ("partial", "ambiguous")
            if any(term in status_low for term in unverified_terms):
                status = "unverified"
            elif any(term in status_low for term in partial_terms):
                status = "partial"
            else:
                status = "verified"
            status_text = status_source or "Source recorded"
            item = {
                "title": title,
                "text": amount or (
                    criteria
                    if not limitations
                    else "Details are recorded in the conditions below."
                ),
                "condition": detail,
                "status": status,
                "statusText": status_text,
                "sources": source_ids,
            }
            if low_category == "add-on":
                add_on.append(item)
            else:
                benefits.append(item)

        annual_fee_text = _text(summary.get("Standard Annual Fee"))
        currency = _text(summary.get("Currency")) or "INR"
        annual_fee = _money_number(annual_fee_text)
        joining_fee = _text(summary.get("Standard Joining Fee"))
        annual_label = _money_label(annual_fee_text, currency)
        apr = _text(summary.get("Purchase APR")) or "Unverified"
        forex = _text(summary.get("Foreign Transaction Fee")) or "Unverified"
        dcc = _text(summary.get("DCC Markup")) or "Unverified"
        cash_fee = _text(summary.get("Cash Advance Fee")) or "Unverified"
        official_url = _text(summary.get("Official Source Link"))
        card_image_url = _text(summary.get("Card Image URL"))
        reward_detail = _text(summary.get("Base Rewards Earn Rate"))
        reward_short = _compact_rewards(reward_detail)
        sources = []
        for source in card_sources:
            source_id = _text(source.get("Source ID"))
            url = _text(source.get("Official URL"))
            if not url:
                continue
            detail = "; ".join(part for part in (
                _text(source.get("Relevant section / page")),
                _text(source.get("Document date / effective date")),
                _text(source.get("Applicability / limitation")),
            ) if part)
            source_title = (
                _text(source.get("Official document/page title"))
                or "Official issuer source"
            )
            sources.append([source_title, detail, url, source_id])

        cashback_rate, cashback_title, cashback_conditions = _cashback_summary(
            card_rules
        )
        category_reward_scores = _category_reward_scores(card_rules)

        guidance_key = "Condition-Based Use Signals (Not a Recommendation)"
        cards.append({
            "id": card_id,
            "name": _text(summary.get("Card Name")),
            "issuer": _text(summary.get("Issuer")),
            "market": _text(summary.get("Market")),
            "network": _text(summary.get("Network")),
            "currency": currency,
            "annualFee": annual_fee,
            "feeLabel": annual_label,
            "joiningFee": joining_fee,
            "apr": _compact_rate(apr),
            "aprDetail": apr,
            "forex": _compact_percent(forex),
            "forexDetail": forex,
            "dcc": _compact_percent(dcc),
            "cashFee": cash_fee[:44],
            "cashFeeDetail": cash_fee,
            "rewardShort": reward_short,
            "rewardDetail": reward_detail,
            "checked": _text(summary.get("Research Checked")),
            "categoryTags": _benefit_tags(card_rules),
            "cashbackRate": cashback_rate,
            "cashbackTitle": cashback_title,
            "cashbackConditions": cashback_conditions,
            "categoryRewardScores": category_reward_scores,
            "fees": fees,
            "eligibility": (
                eligibility
                or _sentences(_text(summary.get("Eligibility")))
                or ["Eligibility details not recorded."]
            ),
            "guidance": _text(summary.get(guidance_key)),
            "caveats": _text(summary.get("Unresolved Terms / Conflicts")),
            "officialUrl": official_url,
            "cardImageUrl": card_image_url,
            "benefits": benefits,
            "addOn": add_on,
            "sources": sources,
        })
    workbook.close()
    return cards


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args: Any, **kwargs: Any) -> None:
        super().__init__(*args, directory=str(WEB_DIR), **kwargs)

    def do_GET(self) -> None:
        request_path = self.path.split("?", 1)[0]
        if request_path == "/api/cards":
            try:
                response_data = {
                    "cards": cards_from_workbook(),
                    "workbook": WORKBOOK.name,
                }
                payload = json.dumps(
                    response_data, ensure_ascii=False
                ).encode("utf-8")
            except Exception as error:
                error_data = {"error": str(error)}
                payload = json.dumps(
                    error_data, ensure_ascii=False
                ).encode("utf-8")
                self.send_response(500)
            else:
                self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
            return
        if request_path == "/download/workbook":
            try:
                payload = WORKBOOK.read_bytes()
            except OSError as error:
                self.send_error(404, str(error))
                return
            self.send_response(200)
            content_type = (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
            self.send_header("Content-Type", content_type)
            disposition = f'attachment; filename="{WORKBOOK.name}"'
            self.send_header("Content-Disposition", disposition)
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
            return
        super().do_GET()

    def log_message(self, format: str, *args: Any) -> None:
        print(f"[{self.log_date_time_string()}] {format % args}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Serve Cardwise from the comparison workbook."
    )
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=12000)
    args = parser.parse_args()
    server = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"Cardwise is running at http://{args.host}:{args.port}/")
    print(f"Workbook source: {WORKBOOK}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping Cardwise.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
