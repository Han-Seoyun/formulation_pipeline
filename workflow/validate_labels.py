"""
Step 5 — 라벨 품질 검증
workflow/label_strata/labels_gold.csv 의 gold 라벨을 검증하고
workflow/diagnostics/label_validation.csv 에 이슈를 기록한다.

실행:
  python workflow/validate_labels.py
"""

import ast
import re
import pandas as pd
from pathlib import Path
from datetime import date

BASE_DIR   = Path(__file__).parent.parent
GOLD_CSV   = BASE_DIR / "workflow" / "label_strata" / "labels_gold.csv"
DIAG_DIR   = BASE_DIR / "workflow" / "diagnostics"
OUT_CSV    = DIAG_DIR / "label_validation.csv"

H_CODE_PAT    = re.compile(r"^H[2-4]\d{2}[A-Za-z]*$")   # H361d/H361f/H360Df 등 변형 허용
VALID_SIGNALS = {"DANGER", "WARNING", "CAUTION", ""}     # CAUTION: EPA 라벨 정상 표현
NULL_SIGNALS  = {"NONE", "UNKNOWN", "N/A"}               # 실질적 빈값 — 별도 플래그


def parse_h_codes(raw) -> list[str]:
    """h_codes 컬럼은 Python 리스트 문자열로 저장됨 — ast.literal_eval 필수."""
    if pd.isna(raw) or str(raw).strip() in {"", "nan", "[]", "null"}:
        return []
    try:
        result = ast.literal_eval(str(raw))
        return [str(c).strip() for c in result] if isinstance(result, list) else []
    except Exception:
        return []


def main():
    if not GOLD_CSV.exists():
        raise FileNotFoundError(
            f"입력 파일 없음: {GOLD_CSV}\n"
            "먼저 python workflow/export_label_strata.py (Step 4) 를 실행하세요."
        )

    df = pd.read_csv(GOLD_CSV, low_memory=False)
    DIAG_DIR.mkdir(parents=True, exist_ok=True)

    issues = []

    for _, row in df.iterrows():
        fid = row.get("Formulation_ID", "UNKNOWN")

        # ── h_codes 검증 ──────────────────────────────────────────────────────
        h_list = parse_h_codes(row.get("h_codes"))
        if not h_list:
            issues.append({"Formulation_ID": fid, "issue": "gold_empty_hcodes",
                           "value": str(row.get("h_codes", ""))})
        else:
            for code in h_list:
                if not H_CODE_PAT.match(code):
                    issues.append({"Formulation_ID": fid,
                                   "issue": f"invalid_hcode",
                                   "value": code})

        # ── signal_word 검증 ──────────────────────────────────────────────────
        sw = str(row.get("signal_word", "")).strip().upper()
        if sw in NULL_SIGNALS:
            issues.append({"Formulation_ID": fid, "issue": "gold_null_signal_word",
                           "value": sw})
        elif sw not in VALID_SIGNALS:
            issues.append({"Formulation_ID": fid, "issue": "invalid_signal_word",
                           "value": sw})

        # ── ghs_hazard_classes 비어있는 gold 행 ───────────────────────────────
        ghs = str(row.get("ghs_hazard_classes", "")).strip()
        if ghs in {"", "nan", "[]", "null"}:
            issues.append({"Formulation_ID": fid, "issue": "gold_empty_ghs_classes",
                           "value": ghs})

    issues_df = pd.DataFrame(issues, columns=["Formulation_ID", "issue", "value"])
    issues_df.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")

    # 이슈 유형별 요약
    total = len(df)
    n_issues = len(issues_df)
    print(f"=== validate_labels 완료 ===")
    print(f"  검증 대상 : {total}행 (gold)")
    print(f"  총 이슈   : {n_issues}건")
    if n_issues:
        for issue_type, cnt in issues_df["issue"].value_counts().items():
            print(f"    {issue_type}: {cnt}건")
    print(f"\n수락 기준:")
    fmt_errors = issues_df[issues_df["issue"] == "invalid_hcode"]
    sig_errors = issues_df[issues_df["issue"] == "invalid_signal_word"]
    print(f"  H-code 형식 오류  : {len(fmt_errors)}건 (목표: 0)")
    print(f"  signal_word 오류  : {len(sig_errors)}건 (목표: 0)")
    print(f"\n  검증 리포트 저장 : {OUT_CSV}")


if __name__ == "__main__":
    main()
