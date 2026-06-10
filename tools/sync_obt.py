# -*- coding: utf-8 -*-
"""Sync the 'Data' sheet of [WarZ] OBT Plan - June.xlsx into warz_data.json.
Adds missing items + updates existing. Column layout has an extra 'Tier' column
(Description is at index 13)."""
import openpyxl, json, sys, io, os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

XLSX = r"C:\Users\almon\OneDrive\Desktop\WARZ WEBSITE PUBLISHED\[WarZ] OBT Plan - June.xlsx"
JSON_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "warz_data.json"))


def cell_text(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return v
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    return str(v).strip()


def norm(v, default="-"):
    s = cell_text(v)
    if s == "" or s is True or s is False:
        return default if s == "" else s
    return s


def parse_items(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Data"]
    rows = list(ws.iter_rows(values_only=True))
    items = []
    for row in rows[3:]:  # header at row index 2
        def g(i):
            return row[i] if len(row) > i else None
        item_id = cell_text(g(2))
        name = cell_text(g(4))
        if not item_id or not name:
            continue
        items.append({
            "id": item_id,
            "name": name,
            "price": norm(g(5)),
            "gcPrice": norm(g(6)),
            "dollarPrice": norm(g(7)),
            "type": cell_text(g(8)),
            "stackable": cell_text(g(9)),
            "image": cell_text(g(3)),
            "description": cell_text(g(13)),
            "confirmed": bool(g(0)) if g(0) is not None else False,
        })
    return items


def main():
    excel_items = parse_items(XLSX)
    print(f"Excel 'Data' items: {len(excel_items)}")

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
            added.append(f'{eid} [{ei["type"]}] {ei["name"]}')
        else:
            cur = by_id[eid]
            changed = [k for k, v in ei.items() if str(cur.get(k, "")) != str(v)]
            if changed:
                cur.update(ei)
                updated.append(f'{eid} {ei["name"]} ({", ".join(changed)})')

    ordered_ids = [str(it["id"]) for it in existing]
    new_ids = [str(ei["id"]) for ei in excel_items if str(ei["id"]) not in ordered_ids]
    final = [by_id[i] for i in ordered_ids] + [by_id[i] for i in new_ids]

    data["items"] = final
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")

    print(f"\n=== ADDED {len(added)} ===")
    for a in added:
        print("  +", a)
    print(f"\n=== UPDATED {len(updated)} (showing first 25) ===")
    for u in updated[:25]:
        print("  ~", u)
    if len(updated) > 25:
        print(f"  ... +{len(updated)-25} more")
    print(f"\nTotal items now: {len(final)}")


if __name__ == "__main__":
    main()
