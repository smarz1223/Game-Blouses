"""
Builds index.html from template.html + data/gameblouses_data.json.
Never edit index.html by hand: change template.html or fetch_data.py and rebuild.
"""
import json

TEMPLATE = "template.html"
DATA = "data/gameblouses_data.json"
OUT = "index.html"
PLACEHOLDER = "/*__DATA__*/null"


def main():
    with open(TEMPLATE, encoding="utf-8") as f:
        html = f.read()
    with open(DATA, encoding="utf-8") as f:
        data = json.load(f)
    if html.count(PLACEHOLDER) != 1:
        raise SystemExit(f"Template must contain exactly one {PLACEHOLDER}")
    blob = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = html.replace(PLACEHOLDER, blob)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Wrote {OUT} ({len(html):,} bytes)")


if __name__ == "__main__":
    main()
