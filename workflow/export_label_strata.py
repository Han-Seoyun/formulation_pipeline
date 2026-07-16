"""
Step 4 — 라벨 계층별 CSV 출력
formulation_characteristics_output.csv의 label_tier 컬럼을 기준으로
gold / combined(gold+silver) 서브셋을 workflow/label_strata/ 에 저장한다.

실행:
  python workflow/export_label_strata.py
"""

import pandas as pd
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
INPUT_CSV = BASE_DIR / "artifacts" / "formulation_characteristics_output.csv"
OUT_DIR   = BASE_DIR / "workflow" / "label_strata"

LABEL_COLS = [
    "Formulation_ID",
    "ghs_hazard_classes",
    "h_codes",
    "signal_word",
    "ghs_hazard_classes_evidence",
    "h_codes_evidence",
    "label_tier",
]


def main():
    if not INPUT_CSV.exists():
        raise FileNotFoundError(
            f"입력 파일 없음: {INPUT_CSV}\n"
            "먼저 python workflow/extract_formulation_features.py 를 실행하세요."
        )

    df = pd.read_csv(INPUT_CSV, low_memory=False)

    if "label_tier" not in df.columns:
        # Step 3 미적용 CSV 대응: evidence 컬럼에서 즉시 도출
        print("  경고: label_tier 컬럼 없음 → evidence 컬럼에서 도출 (Step 3 적용 권장)")
        def _derive_tier(row) -> str:
            if (row.get("h_codes_evidence") == "pdf_regex" or
                    row.get("ghs_hazard_classes_evidence") == "pdf_regex"):
                return "gold"
            if (row.get("h_codes_evidence") == "value_no_pdf" or
                    row.get("ghs_hazard_classes_evidence") == "value_no_pdf"):
                return "silver"
            return "none"
        df["label_tier"] = df.apply(_derive_tier, axis=1)

    # 존재하는 컬럼만 선택 (컬럼 없을 경우 방어)
    keep = [c for c in LABEL_COLS if c in df.columns]
    missing = [c for c in LABEL_COLS if c not in df.columns]
    if missing:
        print(f"  경고: 다음 컬럼 없음 (출력 생략): {missing}")

    OUT_DIR.mkdir(parents=True, exist_ok=True)

    df_gold     = df[df["label_tier"] == "gold"][keep]
    df_combined = df[df["label_tier"].isin(["gold", "silver"])][keep]

    gold_path     = OUT_DIR / "labels_gold.csv"
    combined_path = OUT_DIR / "labels_combined.csv"

    df_gold.to_csv(gold_path,     index=False, encoding="utf-8-sig")
    df_combined.to_csv(combined_path, index=False, encoding="utf-8-sig")

    print(f"=== export_label_strata 완료 ===")
    print(f"  gold     : {len(df_gold):,}행  → {gold_path}")
    print(f"  combined : {len(df_combined):,}행  → {combined_path}")

    # 수락 기준 체크
    gold_nulls = df_gold["ghs_hazard_classes"].apply(
        lambda v: str(v).strip() in {"", "nan", "[]", "null"}
    ).sum() if "ghs_hazard_classes" in df_gold.columns else 0

    print(f"\n수락 기준:")
    print(f"  label_tier null 수  : {df['label_tier'].isna().sum()} (목표: 0)")
    print(f"  gold ghs_null 수    : {gold_nulls} (목표: 0)")
    print(f"  gold 행 수          : {len(df_gold)} (현재 기준선 ≥226)")
    print(f"  combined 행 수      : {len(df_combined)} (현재 기준선 ≥556)")


if __name__ == "__main__":
    main()
