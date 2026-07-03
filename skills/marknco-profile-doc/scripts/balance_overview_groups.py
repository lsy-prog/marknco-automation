#!/usr/bin/env python3
"""
좌우로 나란히 배치된(overviewGroups) 오버뷰 표의 높이를 실측 기반으로 정밀하게 맞추는 스크립트.

배경
----
overviewGroups(2단 나란히 배치)를 쓰면, 필드 값 길이가 회사마다 달라서
특정 회사는 성과키워드/홈페이지 같은 필드가 줄바꿈되어 한쪽 표만 길어지는 경우가 생긴다.
글자 수 기반으로 줄바꿈 여부를 미리 예측하는 방식(East Asian Width 등)은 근사치라
안전하지 않다 — 실제로 렌더링해서 "표 테두리 좌표"를 직접 재는 것이 유일하게 정확한 방법이다.

동작 방식 (2-pass)
----
1. config.json(overviewGroups) 그대로 1차 생성 → PDF 변환.
2. 각 PDF에서 좌/우 표의 실제 테두리 하단선 y좌표를 측정해서 차이(diff, pt)를 구한다.
   (텍스트 라벨 위치가 아니라 "선(line)" 좌표를 봐야 정확하다 — 텍스트는 셀 안에서
   중앙정렬되므로 셀이 커져도 텍스트 위치는 크게 안 변할 수 있음)
3. diff가 있는 회사는 companies.json에 `__overviewFillerTwips` 필드를 주입한다.
   → make_profile_docs.js가 이 필드를 보고 부족한 쪽 그룹의 "모든 행에 균등하게"
     여백을 나눠서 셀을 살짝씩 키운다 (빈 행을 추가하지 않음 — 사용자 요구사항).
4. 2차 재생성 → 재측정으로 수렴 확인 (보통 0.5pt 이내로 수렴, 오차는 PDF 렌더링
   반올림 수준이라 육안으로 안 보임).

전제조건
----
- config.json에 "overviewGroups"가 있어야 함 (2개 그룹, 좌/우 나란히 배치 구조).
- 각 그룹의 "마지막 행 label"이 회사마다 동일해야 함 (표 구성이 통일되어 있어야
  측정 기준점(anchor)으로 쓸 수 있음). 예: 왼쪽 그룹 마지막 행 "최신고용인원",
  오른쪽 그룹 마지막 행 "2025년 매출" 같은 식.

사용법
----
python3 balance_overview_groups.py \\
  --companies companies.json \\
  --config config.json \\
  --docs-dir out \\
  --script /path/to/make_profile_docs.js \\
  [--tolerance-pt 1.0] [--max-passes 2]

companies.json은 in-place로 업데이트되고, docs-dir 안의 docx/pdf도 보정 반영해서
덮어쓴다. 완료 후 별도 병합(qpdf 등)만 하면 됨.
"""
import argparse
import json
import subprocess
import sys
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    print("pdfplumber가 필요합니다: pip install pdfplumber --break-system-packages", file=sys.stderr)
    sys.exit(1)


def find_group_anchor_labels(config_path):
    """config.json의 overviewGroups에서 각 그룹의 마지막 행 label을 anchor로 추출."""
    cfg = json.loads(Path(config_path).read_text(encoding="utf-8"))
    groups = cfg.get("overviewGroups")
    if not groups or len(groups) < 2:
        raise ValueError("config.json에 overviewGroups가 2개 이상 없습니다. 이 스크립트는 좌우 2단 배치 구조 전용입니다.")
    return [g["rows"][-1]["label"] for g in groups]


def measure_bottom_lines(pdf_path, anchor_x_split=300, top_bound=340):
    """PDF 1페이지에서 좌/우 표의 테두리 하단선 y좌표(top, points)를 각각 반환.
    anchor_x_split: 좌/우 표를 구분하는 x좌표 기준(pt). 기본 여백(1080 DXA≒54pt) + 표 절반너비 기준 300pt 근처.
    top_bound: 오버뷰 표가 이 y좌표 아래로는 없다고 가정하는 컷오프(pt). 본문 섹션 구분선과 섞이지 않도록.
    """
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        left = [round(l["top"], 1) for l in page.lines
                if abs(l["x0"] - l["x1"]) > 5 and l["x0"] < anchor_x_split and l["x1"] < anchor_x_split + 10 and l["top"] < top_bound]
        right = [round(l["top"], 1) for l in page.lines
                 if abs(l["x0"] - l["x1"]) > 5 and l["x0"] > anchor_x_split and l["top"] < top_bound]
        left_bottom = max(left) if left else None
        right_bottom = max(right) if right else None
        return left_bottom, right_bottom


def run_generation(script, companies_path, config_path, out_dir):
    cmd = ["node", str(script), "--companies", str(companies_path), "--config", str(config_path), "--out", str(out_dir)]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        print(r.stdout, r.stderr, file=sys.stderr)
        raise RuntimeError("make_profile_docs.js 실행 실패")


def convert_to_pdf(docx_path):
    # 이 저장소 안의 공용 soffice 래퍼를 사용 (없으면 soffice 직접 호출로 대체)
    # 반드시 docx가 있는 폴더를 cwd로 지정해야 PDF가 같은 폴더에 생성됨
    wrapper = Path("/mnt/skills/public/docx/scripts/office/soffice.py")
    if wrapper.exists():
        cmd = ["python3", str(wrapper), "--headless", "--convert-to", "pdf", docx_path.name]
    else:
        cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(docx_path.parent), docx_path.name]
    subprocess.run(cmd, capture_output=True, text=True, cwd=str(docx_path.parent))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--companies", required=True)
    ap.add_argument("--config", required=True)
    ap.add_argument("--docs-dir", required=True, help="docx/pdf가 생성되는(생성될) 폴더")
    ap.add_argument("--script", required=True, help="make_profile_docs.js 경로")
    ap.add_argument("--name-field", default="기업명")
    ap.add_argument("--tolerance-pt", type=float, default=1.0)
    ap.add_argument("--max-passes", type=int, default=2)
    ap.add_argument("--skip-first-generate", action="store_true", help="docs-dir에 이미 1차 생성물이 있으면 생성 단계 생략")
    args = ap.parse_args()

    companies_path = Path(args.companies)
    config_path = Path(args.config)
    docs_dir = Path(args.docs_dir)

    cfg = json.loads(config_path.read_text(encoding="utf-8"))
    prefix = cfg.get("outputPrefix", "profile")
    anchor_labels = find_group_anchor_labels(config_path)
    print(f"기준 라벨(그룹별 마지막 행): {anchor_labels}")

    for pass_no in range(1, args.max_passes + 1):
        if not (pass_no == 1 and args.skip_first_generate):
            print(f"[pass {pass_no}] 생성 중...")
            run_generation(args.script, companies_path, config_path, docs_dir)
            for docx in docs_dir.glob("*.docx"):
                convert_to_pdf(docx)

        companies = json.loads(companies_path.read_text(encoding="utf-8"))
        need_fix = 0
        for c in companies:
            name = c[args.name_field]
            pdf_path = docs_dir / f"{prefix}_{name}.pdf"
            if not pdf_path.exists():
                print(f"  경고: {pdf_path} 없음, 건너뜀")
                continue
            left_bottom, right_bottom = measure_bottom_lines(pdf_path)
            if left_bottom is None or right_bottom is None:
                continue
            diff_pt = round(left_bottom - right_bottom, 2)  # 양수 = 왼쪽이 더 김
            if abs(diff_pt) <= args.tolerance_pt:
                # 이미 정렬된 상태 — 기존 보정값이 있어서 정렬된 것일 수 있으므로 절대 건드리지 않음
                # (여기서 pop 하면 "보정 덕분에 맞은 것"과 "원래 안 맞아도 됐던 것"을 구분 못 하고
                #  기존 보정값을 지워버려서 다음 생성 때 다시 어긋나는 버그가 있었음 — 재발 방지)
                continue
            need_fix += 1
            twips = round(abs(diff_pt) * 20)
            group_idx = 1 if diff_pt > 0 else 0  # 짧은 쪽 그룹에 보정 적용
            c["__overviewFillerTwips"] = {str(group_idx): twips}

        companies_path.write_text(json.dumps(companies, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        print(f"[pass {pass_no}] 보정 필요: {need_fix}개 (허용오차 {args.tolerance_pt}pt)")

        if need_fix == 0:
            print("✅ 모든 회사 좌우 표 높이 정렬 완료.")
            return

    print(f"⚠️ {args.max_passes}회 보정 후에도 일부 남아있을 수 있습니다. 위 로그에서 need_fix 값을 확인하세요.")


if __name__ == "__main__":
    main()
