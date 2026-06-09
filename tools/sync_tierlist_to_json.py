# -*- coding: utf-8 -*-
"""Sync the 'Data' sheet of Tierlist.xlsx into warz_data.json.
Adds missing items and updates existing ones (Excel is source of truth for
fields), preserving any JSON-only items. Reports a summary."""
import openpyxl, json, sys, io, os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

XLSX = r"D:\Download\Tierlist.xlsx"
JSON_PATH = os.path.join(os.path.dirname(__file__), "..", "warz_data.json")
JSON_PATH = os.path.abspath(JSON_PATH)


def cell_text(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return v
    if isinstance(v, float):
        # Drop trailing .0 for whole numbers
        if v.is_integer():
            return str(int(v))
        return str(v)
    return str(v).strip()


def norm(v, default="-"):
    s = cell_text(v)
    if s == "" or s is True or s is False:
        return default if s == "" else s
    return s


def parse_excel_items(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Data"]
    items = []
    rows = list(ws.iter_rows(values_only=True))
    # Header is at row index 2 (0-based); data from row 3
    for row in rows[3:]:
        item_id = cell_text(row[2]) if len(row) > 2 else ""
        if not item_id:
            continue
        name = cell_text(row[4]) if len(row) > 4 else ""
        if not name:
            continue
        confirmed = bool(row[0]) if len(row) > 0 and row[0] is not None else False
        items.append({
            "id": item_id,
            "name": name,
            "price": norm(row[5]) if len(row) > 5 else "-",
            "gcPrice": norm(row[6]) if len(row) > 6 else "-",
            "dollarPrice": norm(row[7]) if len(row) > 7 else "-",
            "type": cell_text(row[8]) if len(row) > 8 else "",
            "stackable": cell_text(row[9]) if len(row) > 9 else "",
            "image": cell_text(row[3]) if len(row) > 3 else "",
            "description": cell_text(row[12]) if len(row) > 12 else "",
            "confirmed": confirmed,
        })
    return items


def main():
    excel_items = parse_excel_items(XLSX)
    print(f"Excel 'Data' sheet items: {len(excel_items)}")

    with open(JSON_PATH, encoding="utf-8") as f:
        data = json.load(f)
    existing = data.get("items", [])
    print(f"Current warz_data.json items: {len(existing)}")

    by_id = {str(it["id"]): it for it in existing}

    added, updated = [], []
    for ei in excel_items:
        eid = str(ei["id"])
        if eid not in by_id:
            by_id[eid] = ei
            added.append(f'{eid} — {ei["name"]}')
        else:
            cur = by_id[eid]
            changed = []
            for k, v in ei.items():
                if str(cur.get(k, "")) != str(v):
                    changed.append(k)
                    cur[k] = v
            if changed:
                updated.append(f'{eid} — {ei["name"]} ({", ".join(changed)})')

    # Preserve original order, append new ones at the end
    ordered_ids = [str(it["id"]) for it in existing]
    new_ids = [str(ei["id"]) for ei in excel_items if str(ei["id"]) not in ordered_ids]
    final = [by_id[i] for i in ordered_ids] + [by_id[i] for i in new_ids]

    data["items"] = final
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")

    print(f"\n=== SUMMARY ===")
    print(f"Added:   {len(added)}")
    for a in added:
        print(f"   + {a}")
    print(f"Updated: {len(updated)}")
    for u in updated[:40]:
        print(f"   ~ {u}")
    if len(updated) > 40:
        print(f"   ... and {len(updated) - 40} more")
    print(f"Total now: {len(final)}")


if __name__ == "__main__":
    main()
