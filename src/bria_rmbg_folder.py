
#!/usr/bin/env python
from __future__ import annotations

from pathlib import Path

from PIL import Image
from transformers import pipeline


def main(input_dir: str, output_dir: str) -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Hugging Face BRIA-RMBG 1.4 파이프라인 로드
    # 모델 카드 공식 예시: pipeline("image-segmentation", model="briaai/RMBG-1.4") :contentReference[oaicite:3]{index=3}
    print("📦 모델 로딩 중... (처음 한 번만 시간 좀 걸립니다)")
    pipe = pipeline(
        "image-segmentation",
        model="briaai/RMBG-1.4",
        trust_remote_code=True,
        device="cpu",   
    )


    exts = {".jpg", ".jpeg", ".png", ".webp"}

    count = 0
    for p in sorted(in_dir.iterdir()):
        if not p.is_file():
            continue
        if p.suffix.lower() not in exts:
            continue

        out_path = out_dir / f"{p.stem}_bria.png"

        print(f"▶ 처리 중: {p.name} -> {out_path.name}")
        # 이미지 열어서 pipeline에 전달
        img = Image.open(p).convert("RGB")
        # pipe()가 바로 배경 제거된 PIL 이미지를 리턴 (공식 예시 참고) :contentReference[oaicite:4]{index=4}
        out_img = pipe(img)
        out_img.save(out_path)
        count += 1

    print(f"✅ 완료: {count}개 이미지 처리 ({in_dir} → {out_dir})")


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 3:
        print("사용법: python bria_rmbg_folder.py <input_dir> <output_dir>")
        sys.exit(1)

    main(sys.argv[1], sys.argv[2])
