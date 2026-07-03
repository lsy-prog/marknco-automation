#!/usr/bin/env python3
"""
K-PATH 접수시트 전용: "정리 시트"(지원서 자체 응답) + "innoclaw 데이터"(혁신의숲 보강 데이터)
를 기업명 기준으로 합쳐서 하나의 companies.json으로 만드는 스크립트.

배경
----
접수시트(예: 2026_KT_KPATH_접수시트.xlsx)는 한 워크북 안에 시트가 여러 개인데,
프로필 문서에 필요한 건 이 중 딱 2개뿐이다:
  - "정리 시트"  : 지원기업이 폼(Typeform 등)으로 직접 제출한 응답 (제품소개, 관심 트랙,
                   사업계획서, 담당자 연락처 등). 여기 A~C열(참고 메모/스크리닝/순번)에
                   "불성실 지원"/"신청 취소" 같은 값이 있으면 그 행은 걸러야 한다.
  - "innoclaw 데이터" : 혁신의숲에서 끌어온 보강 데이터. 컬럼 구성이 marknco-profile-doc
                   스킬(롱리스트용)에서 쓰던 "전체 후보" 시트와 100% 동일하다 —
                   AI요약/보유기술/재무(자산·부채·자본)/매출/영업이익/순이익/월별 고용 등.

이 스크립트가 하는 일
----
1. extract_data.py로 "정리 시트"를 추출한다 (skip-keywords="취소,불성실" 필수 — 스크리닝 컬럼이
   A~C열 안에 있어서 자동으로 걸러짐. 이 단계 생략하면 취소/불성실 지원 건이 그대로 섞여 들어간다).
2. extract_data.py로 "innoclaw 데이터"를 추출한다 (skip-keywords 없음 — 이 시트엔 스크리닝
   컬럼이 없어서 필터링 의미가 없다. 1번에서 이미 걸러진 기업명 목록으로 나중에 맞춰짐).
3. 기업명 기준으로 두 딕셔너리를 merge — 정리 시트 필드가 기본, innoclaw 필드가 덮어씀
   (이름이 겹치는 필드는 없으므로 실제로는 합집합이 된다).
4. URL 정규화, 최신고용인원 단위, 부채/자본 3개년 성장률(자산·매출·영업이익·순이익은 엑셀에
   이미 공식 계산값이 있으므로 재계산하지 않음 — marknco-profile-doc 스킬과 동일한 원칙),
   월별 고용 요약 통계까지 한 번에 처리해서 저장한다.

사용법
----
python3 merge_applicant_data.py \\
  --xlsx 2026_KT_KPATH_접수시트.xlsx \\
  --form-sheet "정리 시트" --form-name-col 4 --form-data-cols "3-31" \\
  --enrich-sheet "innoclaw 데이터" --enrich-name-col 3 --enrich-data-cols "2-106" \\
  --output companies.json

컬럼 번호(--form-data-cols 등)는 시트 구조가 바뀌면 같이 바뀌어야 한다. 실행 전에 항상
엑셀을 직접 열어서 헤더 행을 확인할 것 (SKILL.md Step 1 참고).
"""
import argparse
import json
import subprocess
import sys
from pathlib import Path

MONTHS_ORDER = [
    "2025-01", "2025-02", "2025-03", "2025-04", "2025-05", "2025-06",
    "2025-07", "2025-08", "2025-09", "2025-10", "2025-11", "2025-12",
    "2026-01", "2026-02", "2026-03", "2026-04", "2026-05",
]


def to_float(v):
    try:
        if v in (None, "", "-"):
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


def g3_abs(v0, v1):
    """기준값이 음수여도 부호가 직관과 맞도록 분모에 절댓값 사용 (marknco-profile-doc 4-2-4 참고)"""
    v0f, v1f = to_float(v0), to_float(v1)
    if v0f is None or v1f is None:
        return None
    if v0f == 0:
        return None
    return (v1f - v0f) / abs(v0f)


def run_extract(script_path, xlsx, sheet, name_col, data_cols, out_path, skip_keywords=""):
    cmd = [
        sys.executable, str(script_path),
        "--xlsx", xlsx, "--sheet", sheet,
        "--name-col", str(name_col), "--data-cols", data_cols,
        "--output", str(out_path), "--skip-keywords", skip_keywords,
    ]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        print(r.stdout, r.stderr, file=sys.stderr)
        raise RuntimeError(f"{sheet} 추출 실패")
    return json.loads(Path(out_path).read_text(encoding="utf-8"))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--form-sheet", default="정리 시트")
    ap.add_argument("--form-name-col", type=int, default=4)
    ap.add_argument("--form-data-cols", default="3-31")
    ap.add_argument("--form-skip-keywords", default="취소,불성실")
    ap.add_argument("--enrich-sheet", default="innoclaw 데이터")
    ap.add_argument("--enrich-name-col", type=int, default=3)
    ap.add_argument("--enrich-data-cols", default="2-106")
    ap.add_argument("--output", required=True)
    args = ap.parse_args()

    script_dir = Path(__file__).parent
    extract_script = script_dir / "extract_data.py"

    form = run_extract(extract_script, args.xlsx, args.form_sheet, args.form_name_col,
                        args.form_data_cols, "_tmp_form.json", args.form_skip_keywords)
    enrich = run_extract(extract_script, args.xlsx, args.enrich_sheet, args.enrich_name_col,
                          args.enrich_data_cols, "_tmp_enrich.json", "")
    enrich_by_name = {c["기업명"]: c for c in enrich}

    merged = []
    missing = []
    for c in form:
        name = c["기업명"]
        extra = enrich_by_name.get(name)
        if extra:
            merged.append({**c, **extra, "__has_innoclaw": True})
        else:
            missing.append(name)
            c["__has_innoclaw"] = False
            merged.append(c)
    if missing:
        print(f"경고: innoclaw 데이터에서 못 찾은 기업 {len(missing)}개 — {missing}", file=sys.stderr)

    # 중복 기업명 체크 (파일명이 기업명 기준이라 중복되면 파일이 덮어써짐 — 반드시 사람이 확인해야 함)
    names = [c["기업명"] for c in merged]
    dups = sorted(set(n for n in names if names.count(n) > 1))
    if dups:
        print(f"⚠️  중복 기업명 {len(dups)}개 발견 — 원본 시트 데이터 정합성 확인 필요: {dups}", file=sys.stderr)

    for c in merged:
        hp = c.get("홈페이지 주소") or c.get("홈페이지")
        if hp and isinstance(hp, str) and not hp.startswith(("http://", "https://")):
            hp = "https://" + hp
        if hp:
            c["홈페이지 주소"] = hp
            c["홈페이지"] = hp

        emp = c.get("최신고용인원")
        if emp not in (None, "", "-"):
            c["최신고용인원"] = f"{emp}명"

        c["부채성장률_3개년"] = g3_abs(c.get("2023년 부채"), c.get("2025년 부채"))
        c["자본성장률_3개년"] = g3_abs(c.get("2023년 자본"), c.get("2025년 자본"))
        if to_float(c.get("2023년 자산")) == 0:
            c["최근 3개년 재무 성장률"] = None

        hc_pts, hires_sum, leaves_sum = [], 0, 0
        for m in MONTHS_ORDER:
            hc = to_float(c.get(f"{m} 근무 인원 수"))
            hi = to_float(c.get(f"{m} 입사자"))
            lv = to_float(c.get(f"{m} 퇴사자"))
            if hc is not None:
                hc_pts.append(hc)
            hires_sum += hi or 0
            leaves_sum += lv or 0
        c["__emp_latest"] = hc_pts[-1] if hc_pts else None
        c["__emp_first"] = hc_pts[0] if hc_pts else None
        c["__emp_hires_sum"] = hires_sum if hc_pts else None
        c["__emp_leaves_sum"] = leaves_sum if hc_pts else None
        c["__emp_net_change"] = (hc_pts[-1] - hc_pts[0]) if len(hc_pts) >= 2 else None

        # 관심 트랙 1~3, 기대사항 1~3은 표/불릿에 넣기 좋게 하나의 필드로 통합
        # (0으로 채워진 빈칸, None 다 걸러내고 실제 값 있는 것만 모음)
        def _clean(v):
            return v if v and str(v) != "0" else None
        tracks = [_clean(c.get(f"관심 트랙 {i}")) for i in (1, 2, 3)]
        tracks = [t for t in tracks if t]
        c["__관심트랙_통합"] = ", ".join(tracks) if tracks else None

        expects = [_clean(c.get(f"기대사항 {i}")) for i in (1, 2, 3)]
        expects = [e for e in expects if e]
        c["__기대사항_통합"] = "\n".join(f"- {e}" for e in expects) if expects else None

    Path(args.output).write_text(
        json.dumps(merged, ensure_ascii=False, indent=2, default=str), encoding="utf-8"
    )
    print(f"✅ {len(merged)}개 지원기업 병합 완료 → {args.output}")


if __name__ == "__main__":
    main()
