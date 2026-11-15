# Cardwise

Cardwise is a local credit-card comparison dashboard for India. It compares issuer-published fees, rewards, benefits, eligibility, exclusions, and qualifying conditions. The workbook is the single source of card data; the dashboard loads it at runtime and does not keep a second hardcoded card list.

## Features

- Search cards and filter by issuer, network, benefit focus, and annual-fee ceiling.
- Rank results by verified terms for the selected benefit focus, or by verified cashback when no focus is selected.
- Review benefit conditions, exclusions, unresolved terms, official sources, and card-specific fees.
- Select cards for a side-by-side comparison.
- Download the current workbook from the dashboard.
- Research and update card records using the workspace Credit Card Info agent.

## Quick Start

Requirements: Python 3.10 or later and pip.

From the repository root, install the dependency and start the local server:

```sh
python3 -m pip install -r web/requirements.txt
python3 web/server.py
```

Open <http://127.0.0.1:12000/>. The server binds to localhost by default. Stop it with `Ctrl+C`.

To use a different port:

```sh
python3 web/server.py --port 8001
```

After updating the workbook, select **Refresh workbook** in the dashboard to reload the latest data.

## Project Layout

```text
.
├── .github/agents/credit-card-info.agent.md  # Research and workbook workflow
├── output/credit-card-comparison.xlsx        # Card data and official source register
└── web/
    ├── index.html                            # Dashboard UI
    ├── server.py                             # Local API and workbook download
    └── requirements.txt                      # Python dependencies
```

The workbook contains three main sheets:

- `Card Summary`: one row per card, stable card ID, summary facts, image provenance, and official product-page link.
- `Benefits & Rules`: detailed rewards, fees, eligibility, conditions, caps, exclusions, verification status, and source IDs.
- `Source Register`: official source URLs, document titles, relevant sections, limitations, and access dates.

## Local Endpoints

- `GET /` serves the dashboard.
- `GET /api/cards` returns summaries, detailed rules, source links, and ranking data derived from the workbook.
- `GET /download/workbook` downloads `output/credit-card-comparison.xlsx`.

## Updating Card Research

Use the **Credit Card Info** agent in VS Code. Provide the card name followed by the official issuer benefits-page URL, for example:

```text
Example Card Name — https://issuer.example/credit-card/benefits
```

The agent verifies facts against official issuer pages and linked terms, then updates the existing card record or adds a new one without duplicating an existing product. Review the workbook's source register and unresolved-terms fields when comparing benefits. If a card name or market is ambiguous and no official URL is supplied, the agent may ask for clarification.

## Data and Limitations

- Card benefits, fees, eligibility, and terms can change. Check the source links and access dates in the workbook and confirm current terms with the issuer before applying or spending.
- Reward points and miles are kept in their stated units unless issuer terms explicitly provide a qualifying, fixed-value redemption for a named merchant or category.
- Missing or ambiguous information is marked unresolved or unverified rather than inferred.
- This project provides public product information for comparison; it is not personalized financial advice and does not promise approval, savings, or rewards.
- Do not add card numbers, CVVs, PINs, account credentials, or other payment secrets to the workbook or repository.

## Before Publishing

Before pushing to a public GitHub repository, review the workbook and commit contents for accuracy and licensing. Card artwork is included only when its product variant and image source or permission are recorded; do not add unlicensed images. Confirm that the workbook contains only public product information and no personal or account data.