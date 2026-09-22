import json
from pathlib import Path

archive_dir = Path(r"C:\Users\user\.gemini\antigravity\chat_archive")
index_file = archive_dir / "chat_index.json"
if index_file.exists():
    data = json.loads(index_file.read_text(encoding="utf-8"))
    items = data if isinstance(data, list) else data.get("conversations", [])
    for item in items:
        title = item.get("title") or item.get("topic") or str(item.get("id"))
        print(f"[{item.get('id')}] {title}")
