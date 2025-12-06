#!/usr/bin/env python
from __future__ import annotations

from pathlib import Path
from typing import Iterable, Set

from PIL import Image
from transformers import pipeline, Pipeline


# ===== BRIA RMBG 파이프라인 =====

_BRIA_PIPELINE: Pipeline | None = None


def get_bria_pipeline() -> Pipeline:
    """
    BRIA RMBG-1.4 파이프라인을 전역으로 하나만 로드해서 재사용.
    """
    global _BRIA_PIPELINE
    if _BRIA_PIPELINE is None:
        print("📦 BRIA RMBG-1.4 모델 로딩 중... (처음 한 번만 시간 조금 걸립니다)")
        _BRIA_PIPELINE = pipeline(
            "image-segmentation",
            model="briaai/RMBG-1.4",
            trust_remote_code=True,
            device="cpu",  # Intel Mac이므로 CPU 사용
        )
    return _BRIA_PIPELINE


# ===== 배경제거 & 합성 유틸 =====

def remove_bg_pil(img: Image.Image) -> Image.Image:
    """
    PIL Image 입력 → 배경 제거된 PIL Image 반환.
    """
    pipe = get_bria_pipeline()
    out_img = pipe(img)  # BRIA 커스텀 pipeline: PIL Image 리턴
    return out_img


def compose_on_background(
    fg: Image.Image,
    bg: Image.Image,
    max_ratio: float = 0.8,
) -> Image.Image:
    """
    fg(누끼 이미지)를 bg(1000x1000 배경) 위에 합성한 새 이미지를 반환.

    - max_ratio: 상품 이미지가 배경의 몇 % 크기까지 차지할지 (0~1)
    """
    bg = bg.convert("RGBA")
    fg = fg.convert("RGBA")

    # 1) 배경 기준으로 "최대 허용 박스" 계산 (예: 1000x1000의 80% → 800x800)
    max_w = int(bg.width * max_ratio)
    max_h = int(bg.height * max_ratio)

    # 2) 현재 상품 이미지 크기
    fg_w, fg_h = fg.width, fg.height

    # 3) 80% 박스 안에 "꽉 차도록" 스케일 (확대/축소 둘 다 허용)
    ratio = min(max_w / fg_w, max_h / fg_h)  # ❗ 1.0 제한 제거

    new_size = (int(fg_w * ratio), int(fg_h * ratio))
    fg_resized = fg.resize(new_size, Image.LANCZOS)

    # 4) 중앙 배치
    x = (bg.width - fg_resized.width) // 2
    y = (bg.height - fg_resized.height) // 2

    # 5) 합성
    out = bg.copy()
    out.alpha_composite(fg_resized, dest=(x, y))
    return out


def process_captured_folder(
    images_dir: Path,
    bg_image_path: Path,
    keep_nobg: bool = True,
) -> int:
    """
    폴더 안의 캡처 이미지들을 다음 순서로 처리:

    1. A열 번호와 같은 파일 이름 x.png (숫자만)만 대상으로 삼는다.
    2. x.png → x_org.png 로 백업 (없을 때만)
    3. x_org.png 에 BRIA로 누끼 제거 → x_nobg.png
    4. 1000x1000 배경 위에 x_nobg.png 합성 → x.png 로 최종 저장

    반환값: 처리한 파일 개수.
    """
    images_dir = Path(images_dir)
    bg_image_path = Path(bg_image_path)

    if not images_dir.exists():
        raise FileNotFoundError(f"이미지 폴더가 없습니다: {images_dir}")
    if not bg_image_path.exists():
        raise FileNotFoundError(f"배경 이미지가 없습니다: {bg_image_path}")

    print(f"📂 이미지 폴더: {images_dir}")
    print(f"🖼  배경 이미지: {bg_image_path}")

    # 배경 이미지는 한 번만 로드
    bg = Image.open(bg_image_path).convert("RGBA")

    # 파이프라인도 미리 로드 (첫 호출에서 로딩)
    pipe = get_bria_pipeline()

    count = 0

    for p in sorted(images_dir.iterdir()):
        # 1) PNG & 파일명 숫자만 (예: 1.png, 23.png)
        if not p.is_file():
            continue
        if p.suffix.lower() not in {".png", ".jpg", ".jpeg"}:
            continue
        if not p.stem.isdigit():
            # 예: 1_1.png, 1_spec.png 등은 스킵
            continue

        stem = p.stem  # "1", "2", ...

        org_path = images_dir / f"{stem}_org.png"
        nobg_path = images_dir / f"{stem}_nobg.png"
        final_path = images_dir / f"{stem}.png"

        print(f"\n▶ 처리 대상: {p.name}")

        # 2) 원본 백업: x.png → x_org.png (이미 있으면 건너뛰기)
        if not org_path.exists():
            print(f"  - 원본 백업: {p.name} → {org_path.name}")
            p.rename(org_path)
        else:
            print(f"  - 원본 백업 이미 존재: {org_path.name}")

        # 3) 누끼 제거: x_org.png → x_nobg.png
        print(f"  - BRIA 누끼 제거: {org_path.name} → {nobg_path.name}")
        img_org = Image.open(org_path).convert("RGB")
        out_nobg = pipe(img_org)
        nobg_path.parent.mkdir(parents=True, exist_ok=True)
        out_nobg.save(nobg_path)

        # 4) 배경 합성: x_nobg.png + 1000x1000 → x.png
        print(f"  - 배경 합성 후 최종 저장: {final_path.name}")
        fg = Image.open(nobg_path).convert("RGBA")
        out_final = compose_on_background(fg, bg, max_ratio=0.9)
        out_final.save(final_path)

        # 5) 필요 없으면 x_nobg.png 삭제 옵션
        if not keep_nobg:
            print(f"  - 중간 파일 삭제: {nobg_path.name}")
            nobg_path.unlink(missing_ok=True)

        count += 1

    print(f"\n✅ 전체 완료: {count}개 파일 처리 ({images_dir})")
    return count


# ===== CLI 진입점 =====

def main():
    """
    사용법:
        python process_captured_bria.py <images_dir> <bg_image_path>

    예:
        python process_captured_bria.py \
            "/Users/jeehoonkim/Desktop/상품사진/image_test" \
            "/Users/jeehoonkim/Desktop/상품사진/bg/product_bg_1000.png"
    """
    import sys

    if len(sys.argv) != 3:
        print("사용법: python process_captured_bria.py <images_dir> <bg_image_path>")
        sys.exit(1)

    images_dir = Path(sys.argv[1])
    bg_image_path = Path(sys.argv[2])

    process_captured_folder(images_dir, bg_image_path, keep_nobg=True)


if __name__ == "__main__":
    main()
