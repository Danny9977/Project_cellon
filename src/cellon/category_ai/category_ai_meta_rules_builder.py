from __future__ import annotations

import argparse
import hashlib
import json
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


# -----------------------------
# Utilities
# -----------------------------
def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s or "")


def stable_hash8(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:8]


def resolve_project_root(start: Path) -> Path:
    """
    Find repo root by locating 'src/cellon'.
    """
    start = start.resolve()
    for base in [start, *start.parents]:
        if (base / "src" / "cellon").exists():
            return base
    return Path.cwd().resolve()


def tokenize_category_key(key: str) -> List[str]:
    """
    Coupang category key example:
      '식품>냉장:냉동식품>햄:소시지:어묵:맛살'
    -> ['식품','냉장','냉동식품','햄','소시지','어묵','맛살']
    """
    key = nfc(key).strip()
    parts = re.split(r"[>:/\\|\s]+", key)
    parts = [p.strip() for p in parts if p and p.strip()]

    out: List[str] = []
    for p in parts:
        p2 = re.sub(r"^[\W_]+|[\W_]+$", "", p).strip()
        if len(p2) < 2:
            continue
        if p2 not in out:
            out.append(p2)
    return out


@dataclass
class TemplateItem:
    key: str
    relative_path: str


def load_index(index_path: Path) -> List[TemplateItem]:
    """
    Expected structure (from your uploaded file):
      {
        "root": ".../assets/crawling_temp/coupang_upload_form",
        "templates": [
          {"key":"...", "relative_path":"14_주방용품/....xlsm", ...},
          ...
        ]
      }
    """
    data = json.loads(index_path.read_text(encoding="utf-8"))
    templates = data.get("templates", [])
    items: List[TemplateItem] = []
    for t in templates:
        key = nfc(str(t.get("key", "")).strip())
        rel = nfc(str(t.get("relative_path", "")).strip())
        if key and rel:
            items.append(TemplateItem(key=key, relative_path=rel))
    return items


def infer_group_from_relative_path(relative_path: str, group_map: Dict[str, str]) -> Optional[str]:
    """
    Use first folder prefix '10_식품', '14_주방용품' -> code '10','14' -> map to group.
    """
    rel = nfc(relative_path).replace("\\", "/")
    first = rel.split("/", 1)[0]
    m = re.match(r"^(\d+)_", first)
    if not m:
        return None
    code = m.group(1)
    return group_map.get(code)


def load_json(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        raise RuntimeError(f"❌ meta json 파싱 실패: {path} ({e})")


def write_json(path: Path, obj: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def ensure_header(meta: Dict[str, Any], group: str) -> None:
    """
    Keep a header block for future validation/versioning.
    Do not overwrite existing.
    """
    if "__meta__" not in meta:
        meta["__meta__"] = {"group": group, "version": 1}
        return

    hdr = meta.get("__meta__")
    if isinstance(hdr, dict):
        hdr.setdefault("group", group)
        hdr.setdefault("version", 1)


def build_label_to_meta_key(meta: Dict[str, Any]) -> Dict[str, str]:
    """
    label -> meta_key map (ignore '__meta__').
    """
    out: Dict[str, str] = {}
    for mk, obj in meta.items():
        if mk == "__meta__":
            continue
        if not isinstance(obj, dict):
            continue
        label = nfc(str(obj.get("label", "")).strip())
        if label:
            out[label] = mk
    return out


def make_meta_key(group: str, label: str, used_keys: set[str]) -> str:
    """
    Stable key to avoid reordering issues:
      meta_{group}_{hash(label)}
    """
    base = f"meta_{group}_{stable_hash8(label)}"
    mk = base
    i = 1
    while mk in used_keys:
        mk = f"{base}_{i}"
        i += 1
    return mk


def merge_preserve_user(existing: Dict[str, Any], generated: Dict[str, Any]) -> Tuple[Dict[str, Any], bool]:
    """
    규칙:
    - 기존 값이 '비어있지 않으면' 절대 덮어쓰지 않음
    - 비어있는 필드만 채움
    - list는 기존이 []일 때만 채움
    """
    changed = False
    merged = dict(existing)

    for k, v in generated.items():
        if k not in merged:
            merged[k] = v
            changed = True
            continue

        cur = merged.get(k)

        # list: fill only if empty
        if isinstance(cur, list) and isinstance(v, list):
            if len(cur) == 0 and len(v) > 0:
                merged[k] = v
                changed = True
            continue

        # string/None: fill only if empty
        if cur is None:
            merged[k] = v
            changed = True
            continue
        if isinstance(cur, str) and cur.strip() == "":
            merged[k] = v
            changed = True
            continue

        # dict: fill missing subkeys only
        if isinstance(cur, dict) and isinstance(v, dict):
            for kk, vv in v.items():
                if kk not in cur:
                    cur[kk] = vv
                    changed = True
            merged[k] = cur
            continue

        # otherwise keep user value
        continue

    return merged, changed


# -----------------------------
# Builder
# -----------------------------
def build_meta_rules(
    index_path: Path,
    meta_dir: Path,
    group: Optional[str] = None,
    group_map: Optional[Dict[str, str]] = None,
    dry_run: bool = False,
    verbose: bool = False,
) -> Dict[str, Dict[str, int]]:
    """
    Return summary: {group: {added, updated, skipped, total}}
    """
    if group_map is None:
        # Danny님의 현재 폴더 구조 기준
        group_map = {
            "14": "kitchen",
            "10": "food",
        }

    items = load_index(index_path)

    # group -> labels(keys)
    group_to_labels: Dict[str, List[str]] = {}
    for it in items:
        g = infer_group_from_relative_path(it.relative_path, group_map)
        if not g:
            continue
        if group and g != group:
            continue
        group_to_labels.setdefault(g, []).append(it.key)

    summary: Dict[str, Dict[str, int]] = {}

    for g, labels in group_to_labels.items():
        meta_path = meta_dir / f"coupang_{g}.json"
        meta = load_json(meta_path)         # Load existing meta rules
        ensure_header(meta, g)              # __meta__ 생성

        used_keys = set(meta.keys())
        label_map = build_label_to_meta_key(meta)

        # de-dup labels (preserve order)
        seen = set()
        uniq: List[str] = []
        for lb in labels:
            lb2 = nfc(lb)
            if lb2 in seen:
                continue
            seen.add(lb2)
            uniq.append(lb2)

        added = updated = skipped = 0

        for label in uniq:
            generated = {
                "label": label,
                "keywords_include": tokenize_category_key(label),
                "keywords_exclude": [],
                "source_costco_paths": [],
                "source_domemae_paths": [],
                "source_owner_paths": [],
            }

            if label in label_map:
                mk = label_map[label]
                existing = meta.get(mk, {})
                if not isinstance(existing, dict):
                    existing = {}
                merged, changed = merge_preserve_user(existing, generated)
                if changed:
                    meta[mk] = merged
                    updated += 1
                    if verbose:
                        print(f"[{g}] updated: {mk}  label='{label}'")
                else:
                    skipped += 1
            else:
                mk = make_meta_key(g, label, used_keys)
                used_keys.add(mk)
                meta[mk] = generated
                added += 1
                if verbose:
                    print(f"[{g}] added: {mk}  label='{label}'")

        summary[g] = {"added": added, "updated": updated, "skipped": skipped, "total": len(uniq)}

        if not dry_run:
            # backup once per run
            if meta_path.exists():
                backup = meta_path.with_suffix(meta_path.suffix + ".bak")
                try:
                    backup.write_text(meta_path.read_text(encoding="utf-8"), encoding="utf-8")
                except Exception:
                    print(f"⚠️ backup failed: {backup}")

            write_json(meta_path, meta)

        if verbose:
            print(f"[{g}] -> {meta_path} (added={added}, updated={updated}, skipped={skipped}, total={len(uniq)})")

    return summary


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build/augment meta rules (meta/coupang_{group}.json) from coupang_upload_index.json",
    )
    parser.add_argument(
        "--index",
        default="assets/cache/coupang_upload_index.json",
        help="Path to coupang_upload_index.json (relative to project root)",
    )
    parser.add_argument(
        "--meta-dir",
        default="src/cellon/rules/meta",
        help="Meta rules directory (relative to project root)",
    )
    parser.add_argument(
        "--group",
        default="",
        help="Restrict to one group (kitchen|food). Empty = all detected groups",
    )
    parser.add_argument("--dry-run", action="store_true", help="Preview only (no file write)")
    parser.add_argument("--verbose", action="store_true", help="Verbose logging")

    args = parser.parse_args()

    root = resolve_project_root(Path(__file__).resolve())
    index_path = (root / nfc(args.index)).resolve()
    meta_dir = (root / nfc(args.meta_dir)).resolve()

    group = nfc(args.group).strip() or None

    if not index_path.exists():
        raise FileNotFoundError(f"index 파일이 없습니다: {index_path}")

    summary = build_meta_rules(
        index_path=index_path,
        meta_dir=meta_dir,
        group=group,
        dry_run=args.dry_run,
        verbose=args.verbose,
    )

    print("✅ category_ai_meta_rules_builder summary")
    if not summary:
        print(" - (no groups detected; index의 relative_path 폴더명이 '10_식품', '14_주방용품' 형태인지 확인하세요)")
        return

    for g, s in summary.items():
        print(f" - {g}: added={s['added']}, updated={s['updated']}, skipped={s['skipped']}, total={s['total']}")


if __name__ == "__main__":
    main()
