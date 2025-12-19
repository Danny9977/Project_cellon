# src/coupang_upload_template_resolver.py

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import json
from typing import Optional

from cellon.config import COUPANG_UPLOAD_FORM_DIR, COUPANG_UPLOAD_INDEX_JSON


@dataclass(frozen=True)
class CoupangTemplateHit:
    key: str
    path: Path                 # 절대경로
    relative_path: str         # root 기준 상대경로
    found_by: str              # "index_json" | "rglob_fallback"


def _load_index_json(index_path: Path) -> dict:
    if not index_path.exists():
        return {}
    try:
        with index_path.open("r", encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        # JSON 파손/인코딩 문제 등 -> 빈 dict로 처리하고 fallback 유도
        return {}


def _resolve_by_index_json(root: Path, template_key: str) -> Optional[CoupangTemplateHit]:
    data = _load_index_json(COUPANG_UPLOAD_INDEX_JSON)
    templates = data.get("templates") or []
    if not templates:
        return None

    # key 일치하는 첫 항목 선택 (중복이 있으면 build 단계에서 경고 출력되도록 함)
    for t in templates:
        if (t.get("key") or "") == template_key:
            rel = t.get("relative_path") or ""
            if not rel:
                return None

            abs_path = (root / rel).resolve()
            if abs_path.exists():
                return CoupangTemplateHit(
                    key=template_key,
                    path=abs_path,
                    relative_path=rel,
                    found_by="index_json",
                )
            # 인덱스에는 있지만 파일이 이동/삭제된 경우 -> None으로 fallback
            return None

    return None


def _resolve_by_rglob(root: Path, template_key: str) -> Optional[CoupangTemplateHit]:
    # A 방식: 재귀탐색 백업
    pattern = f"sellertool_upload_{template_key}.xlsm"
    for p in root.rglob(pattern):
        rel = str(p.relative_to(root))
        return CoupangTemplateHit(
            key=template_key,
            path=p.resolve(),
            relative_path=rel,
            found_by="rglob_fallback",
        )
    return None


def resolve_coupang_upload_template(
    template_key: str,
    *,
    root: Path = COUPANG_UPLOAD_FORM_DIR,
    prefer_index_json: bool = True,
    auto_rebuild_index_if_missing: bool = False,
) -> CoupangTemplateHit:
    """
    템플릿 경로를 찾는다.

    - (B 추천) 인덱스 JSON 기반으로 먼저 찾음
    - 실패 시 (A 백업) rglob() 재귀 탐색으로 찾음

    참고 문구 표기 시에는:
    "coupang_upload_form 내의 쿠팡 템플릿 구조" 같이 폴더 범위를 명시하는 문구를 사용하세요.
    """
    if not root.exists():
        raise FileNotFoundError(f"coupang_upload_form 폴더가 없습니다: {root}")

    # 1) B 우선
    if prefer_index_json:
        hit = _resolve_by_index_json(root, template_key)
        if hit:
            return hit

        # (선택) 인덱스가 아예 없으면 자동 생성 시도
        if auto_rebuild_index_if_missing and (not COUPANG_UPLOAD_INDEX_JSON.exists()):
            # 여기서 import는 순환참조 방지용 로컬 import
            from src.build_coupang_upload_index import build_coupang_upload_index  # noqa
            build_coupang_upload_index()
            hit = _resolve_by_index_json(root, template_key)
            if hit:
                return hit

    # 2) A 백업
    hit = _resolve_by_rglob(root, template_key)
    if hit:
        return hit

    # 3) 최종 실패
    raise FileNotFoundError(
        f"템플릿을 찾지 못했습니다: key='{template_key}'\n"
        f"- (B) index_json: {COUPANG_UPLOAD_INDEX_JSON} 에서 키를 찾지 못했거나 파일이 존재하지 않음\n"
        f"- (A) rglob: {root} 하위에서 sellertool_upload_{template_key}.xlsm 를 찾지 못함\n"
    )
