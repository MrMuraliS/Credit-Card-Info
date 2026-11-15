# Cardwise local dashboard

The dashboard reads its card data directly from `output/credit-card-comparison.xlsx` through the local server. It does not keep a separate card dataset in the HTML.

## Run

From the workspace root:

```sh
python3 -m pip install -r web/requirements.txt
python3 web/server.py
```

Open the URL printed by the server (default `http://127.0.0.1:12000/`). After changing the workbook, use **Refresh workbook** in the page to reload card summaries, benefits and rules, and official sources.

To use a different local port:

```sh
python3 web/server.py --port 8001
```
