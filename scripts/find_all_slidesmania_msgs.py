# -*- coding: utf-8 -*-
import sys
import os
import base64
import struct
import json

cat_path = r"C:\Users\user\Desktop\Ish\Soff.uz\telegram_bot\templates_catalog.json"
with open(cat_path, "r", encoding="utf-8") as f:
    cat = json.load(f)

# Find all templates from SlidesMania or SlidesCarnival
target_templates = [v for v in cat.values() if "SlidesMania" in v.get("local_path", "") or "SlidesCarnival" in v.get("local_path", "")]
print(f"Total target templates: {len(target_templates)}")

results = []
for item in target_templates:
    fid = item.get("pptx_file_id", "")
    if fid:
        try:
            data = base64.urlsafe_b64decode(fid + "===")
            doc_mid = struct.unpack(">H", data[18:20])[0]
            photo_mid = doc_mid - 1
            results.append({
                "id": item["id"],
                "filename": item["filename"],
                "local_path": item["local_path"],
                "doc_mid": doc_mid,
                "photo_mid": photo_mid,
                "thumb_file_id": item.get("thumb_file_id")
            })
        except Exception:
            pass

print(f"Successfully decoded message IDs for {len(results)} templates.")
# Print first 10 and last 10
for r in results[:10]:
    print(f"  ID: {r['id']} | Doc: {r['doc_mid']}, Photo: {r['photo_mid']} | {r['filename'][:40]}")

print("...\nLast 10:")
for r in results[-10:]:
    print(f"  ID: {r['id']} | Doc: {r['doc_mid']}, Photo: {r['photo_mid']} | {r['filename'][:40]}")

with open(r"C:\Users\user\Desktop\Ish\Soff.uz\telegram_bot\scripts\slidesmania_decoded_msgs.json", "w", encoding="utf-8") as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
