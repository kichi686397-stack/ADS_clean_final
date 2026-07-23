# -*- coding: utf-8 -*-
"""Import a pasted Graph Explorer leadgen_forms response into the local library."""
import json
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

import app


def main(source_path, page_id="1076514822219251"):
    source = Path(source_path)
    text = source.read_text(encoding="utf-8")
    pattern = re.compile(
        r'\{\s*"id"\s*:\s*"\[(\d+)\]\([^)]*\)"\s*,\s*'
        r'"name"\s*:\s*"([^"]+)"\s*,\s*'
        r'"status"\s*:\s*"([^"]+)"\s*,\s*'
        r'"locale"\s*:\s*"([^"]+)"\s*,\s*'
        r'"created_time"\s*:\s*"([^"]+)"\s*\}',
        re.S,
    )
    items = []
    for form_id, name, status, locale, created_time in pattern.findall(text):
        parsed = app.parse_form_name(name)
        if not parsed:
            continue
        items.append({
            "form_id": form_id,
            "name": name,
            "status": status,
            "created_time": created_time,
            "locale": locale,
            "page_id": page_id,
            "product": parsed["product"],
            "lang": parsed["lang"],
        })
    if not items:
        raise RuntimeError("没有从导出文本中识别到表单。")

    library_path = app.FORMS_LIBRARY_FILE
    raw = json.loads(library_path.read_text(encoding="utf-8")) if library_path.exists() else {}
    if not isinstance(raw, dict):
        raw = {}
    backup = library_path.with_name(f"forms_library_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
    if library_path.exists():
        shutil.copy2(library_path, backup)

    page = raw.setdefault(page_id, {})
    changed = 0
    for item in items:
        product = item.pop("product")
        lang = item.pop("lang")
        old = (page.get(product) or {}).get(lang) or {}
        if str(item.get("created_time") or "") >= str(old.get("created_time") or ""):
            page.setdefault(product, {})[lang] = item
            if str(old.get("form_id") or old.get("id") or "") != item["form_id"]:
                changed += 1

    library_path.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")
    normalized = app.load_forms_library()
    page_products = ((normalized.get("by_page") or {}).get(page_id) or {}).get("by_product") or {}
    print(json.dumps({
        "recognized": len(items),
        "changed": changed,
        "page_id": page_id,
        "products": len(page_products),
        "backup": str(backup),
        "library": str(library_path),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "1076514822219251")
