#!/usr/bin/env python3
"""
기업별 "성장 데이터" 차트(고용 시계열, 재무 그룹막대, 매출·손익 그룹막대)를 일괄 생성하는 스크립트.

배경 / 언제 쓰는가
----
2페이지 이상으로 "고용 현황 / 재무 현황(자산·부채·자본) / 매출·손익 현황(매출액·영업이익·순이익)"
같은 성장 데이터 섹션을 넣을 때 사용한다 (make_profile_docs.js의 cfg.dataTables + chartField 조합).
1페이지짜리 기본 프로필에는 필요 없다.

이 스크립트가 하는 일
----
1. companies.json을 읽어 각 기업마다 PNG 차트를 그려서 저장한다.
2. 각 기업 레코드에 차트 파일 경로를 `__chart_employment` / `__chart_financial` / `__chart_revenue`
   필드로 추가해서 companies.json에 다시 저장한다.
   → make_profile_docs.js 쪽 config에서 `"chartField": "__chart_financial"` 식으로 참조하면
     기업마다 자기 데이터로 그려진 차트가 자동으로 매칭된다 (하드코딩 이미지 절대 금지 —
     기업이 여러 개면 반드시 chartField 방식을 쓸 것. 정적 이미지를 config에 박아두면
     다른 회사 문서에 엉뚱한 차트가 뜨는 사고가 실제로 있었음).
3. 데이터가 부족한(연도 2개 미만) 기업은 해당 차트를 만들지 않고 필드를 null로 둔다.
   → make_profile_docs.js가 자동으로 차트를 스킵하고 표만 보여준다 (정상 동작, 에러 아님).

전제조건: 엑셀 시트에 아래 컬럼들이 정확히 이 이름으로 있어야 한다. 다르면 --field-map 이나
직접 이 파일을 복사해서 컬럼명 상수를 고쳐서 쓸 것.
- 자산/부채/자본: "{연도}년 자산" / "{연도}년 부채" / "{연도}년 자본" (연도=2023,2024,2025)
- 매출/손익: "{연도}년 매출(원)" / "{연도} 영업이익" / "{연도} 순이익"
- 고용(월별): "{YYYY-MM} 근무 인원 수" / "{YYYY-MM} 입사자" / "{YYYY-MM} 퇴사자"

사용법
----
python3 generate_growth_charts.py \\
  --companies companies.json \\
  --out-dir /home/claude/work/charts \\
  [--months 2025-01:2026-05] [--years 2023,2024,2025] \\
  [--skip employment,financial,revenue 중 생략할 것]

주의: companies.json은 in-place로 업데이트된다 (차트 경로 필드 추가).
"""
import argparse
import json
import os
import re

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np

FONT_CANDIDATES = ["/root/.fonts/Pretendard-Regular.otf", "/root/.fonts/Pretendard-Bold.otf"]

# ── 디자인 시스템 색상 (design-system.md 참고) ──────────────────────────
NAVY = "#1F3B57"
TEAL = "#2E9E8F"
CORAL = "#D96C5F"
GRID = "#D7DCE3"
SUBTLE = "#6B7280"
BODY = "#26282B"
YELLOW_GREEN = "#8BC34A"   # 단일 지표(매출만 등) 차트용 대체 색상
MINT = "#10B981"           # 고용 시계열 - 입사자
GRAY_BAR = "#D7DCE3"       # 고용 시계열 - 퇴사자
BLUE_LINE = "#6C7FE0"      # 고용 시계열 - 근무 인원수 선
REV_TEAL = "#1ABC9C"       # 매출·손익 - 매출액
OPINC_BLUE = "#7B8FE8"     # 매출·손익 - 영업이익
NETINC_GOLD = "#F2C230"    # 매출·손익 - 순이익


def setup_font():
    found = False
    for f in FONT_CANDIDATES:
        if os.path.exists(f):
            fm.fontManager.addfont(f)
            found = True
    plt.rcParams["font.family"] = "Pretendard" if found else plt.rcParams["font.family"]
    plt.rcParams["axes.unicode_minus"] = False
    if not found:
        print("경고: Pretendard 폰트를 찾지 못함 — 기본 폰트로 렌더링됨 (한글이 깨질 수 있음)")


def safe_filename(name):
    return re.sub(r"[^\w가-힣.-]", "_", name)


def to_float(v):
    try:
        if v in (None, "", "-"):
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


# ── 그룹 막대 차트 공통 (자산/부채/자본, 매출/영업이익/순이익 등 3개 지표 x N개년) ──
def make_grouped_bar_chart(name, suffix, years_vals, series_labels, colors, out_dir, ylabel="억원", height=2.3):
    """years_vals: [(period_label, v1_or_None, v2_or_None, v3_or_None), ...]"""
    pts = [(y, a, b, c) for y, a, b, c in years_vals if not (a is None and b is None and c is None)]
    if len(pts) < 2:
        return None
    labels = [p[0] for p in pts]
    s1 = [p[1] if p[1] is not None else 0 for p in pts]
    s2 = [p[2] if p[2] is not None else 0 for p in pts]
    s3 = [p[3] if p[3] is not None else 0 for p in pts]

    x = np.arange(len(labels)); w = 0.25
    fig, ax = plt.subplots(figsize=(9.6, height), dpi=200)
    ax.bar(x - w, s1, width=w, color=colors[0], label=series_labels[0], zorder=3)
    ax.bar(x, s2, width=w, color=colors[1], label=series_labels[1], zorder=3)
    ax.bar(x + w, s3, width=w, color=colors[2], label=series_labels[2], zorder=3)
    ax.set_xticks(x); ax.set_xticklabels(labels, fontsize=10.5)
    ax.set_ylabel(ylabel, fontsize=10, color=SUBTLE)
    ax.axhline(0, color=SUBTLE, linewidth=0.8)
    ax.grid(axis="y", color=GRID, linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    ax.spines["left"].set_color(GRID); ax.spines["bottom"].set_color(GRID)
    # 범례는 반드시 그래프 바깥 위쪽에 — plot 안(예: upper right)에 두면 막대가 높은 기업에서 겹쳐서 깨짐
    ax.legend(loc="lower center", bbox_to_anchor=(0.5, 1.02), ncol=3, frameon=False, fontsize=9.5)

    allvals = s1 + s2 + s3
    vmax, vmin = max(allvals), min(allvals)
    pad = max(abs(vmax), abs(vmin)) * 0.22 or 1
    # 음수 막대 라벨이 잘리지 않도록 y축 여유를 명시적으로 확보할 것 (자동 스케일만 믿으면 라벨이 잘림)
    ax.set_ylim(min(0, vmin) - pad, max(0, vmax) + pad * 0.6)
    for xi, vals in [(x - w, s1), (x, s2), (x + w, s3)]:
        for xx, v in zip(xi, vals):
            va = "bottom" if v >= 0 else "top"  # va를 annotate에 반드시 넘길 것 (누락하면 라벨이 막대에 겹침)
            off = 3 if v >= 0 else -3
            ax.annotate(f"{v:.1f}", (xx, v), textcoords="offset points", xytext=(0, off),
                        ha="center", va=va, fontsize=8, color=BODY)

    plt.tight_layout()
    path = os.path.join(out_dir, f"{safe_filename(name)}_{suffix}.png")
    plt.savefig(path, facecolor="white", bbox_inches="tight")
    plt.close()
    return path


def make_employment_timeseries_chart(name, c, months, out_dir):
    headcount, hires, leaves, x_labels = [], [], [], []
    for m in months:
        hc = to_float(c.get(f"{m} 근무 인원 수"))
        hi = to_float(c.get(f"{m} 입사자"))
        lv = to_float(c.get(f"{m} 퇴사자"))
        if hc is None and hi is None and lv is None:
            continue
        x_labels.append(m)
        headcount.append(hc)
        hires.append(hi if hi is not None else 0)
        leaves.append(lv if lv is not None else 0)

    if len([h for h in headcount if h is not None]) < 2:
        return None

    x = np.arange(len(x_labels))
    fig, ax1 = plt.subplots(figsize=(9.8, 2.15), dpi=200)

    w = 0.36
    ax1.bar(x - w / 2, hires, width=w, color=MINT, zorder=3, label="입사자")
    ax1.bar(x + w / 2, leaves, width=w, color=GRAY_BAR, zorder=3, label="퇴사자")
    ax1.set_ylabel("입/퇴사자 (명)", fontsize=9.5, color=SUBTLE)
    max_flow = max(hires + leaves) if (hires or leaves) else 1
    ax1.set_ylim(0, max_flow * 2.6 or 1)
    ax1.grid(axis="y", color=GRID, linewidth=0.7, zorder=0)
    ax1.set_axisbelow(True)
    ax1.spines["top"].set_visible(False)
    ax1.spines["left"].set_color(GRID); ax1.spines["bottom"].set_color(GRID); ax1.spines["right"].set_color(GRID)
    ax1.tick_params(labelsize=8.5, axis="y")

    ax1.set_xticks(x)
    step = 2 if len(x_labels) > 10 else 1  # 라벨이 붐비지 않도록 격월 표시 (10개 초과 시)
    ax1.set_xticklabels([m if i % step == 0 else "" for i, m in enumerate(x_labels)], fontsize=8.3)

    ax2 = ax1.twinx()
    hc_x = [xi for xi, h in zip(x, headcount) if h is not None]
    hc_y = [h for h in headcount if h is not None]
    ax2.plot(hc_x, hc_y, marker="o", color=BLUE_LINE, linewidth=2.0, markersize=4.5, zorder=4)
    ax2.set_ylabel("근무 인원수 (명)", fontsize=9.5, color=SUBTLE)  # 우측 축 텍스트는 회색 유지 (파란색으로 하면 눈에 튐)
    ax2.set_ylim(0, max(hc_y) * 1.35)
    ax2.tick_params(labelsize=8.5, axis="y", labelcolor=SUBTLE)
    ax2.spines["top"].set_visible(False)

    lines1, labels1 = ax1.get_legend_handles_labels()
    ax2.plot([], [], marker="o", color=BLUE_LINE, linewidth=2.0, markersize=4.5, label="근무 인원수")
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="lower center", bbox_to_anchor=(0.5, 1.02),
               ncol=3, frameon=False, fontsize=9)

    plt.tight_layout()
    path = os.path.join(out_dir, f"{safe_filename(name)}_employment_ts.png")
    plt.savefig(path, facecolor="white", bbox_inches="tight")
    plt.close()
    return path


def month_range(start, end):
    """'2025-01' ~ '2026-05' 같은 범위를 월 리스트로 전개"""
    sy, sm = map(int, start.split("-"))
    ey, em = map(int, end.split("-"))
    months = []
    y, m = sy, sm
    while (y, m) <= (ey, em):
        months.append(f"{y:04d}-{m:02d}")
        m += 1
        if m > 12:
            m = 1; y += 1
    return months


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--companies", required=True)
    ap.add_argument("--out-dir", required=True, help="차트 PNG를 저장할 폴더")
    ap.add_argument("--name-field", default="기업명")
    ap.add_argument("--years", default="2023,2024,2025", help="재무/매출·손익 차트에 쓸 연도 (쉼표 구분)")
    ap.add_argument("--months", default="2025-01:2026-05", help="고용 시계열 차트 기간 (시작:끝, YYYY-MM)")
    ap.add_argument("--skip", default="", help="생략할 차트 종류 (예: employment,financial,revenue)")
    args = ap.parse_args()

    setup_font()
    os.makedirs(args.out_dir, exist_ok=True)
    years = args.years.split(",")
    m_start, m_end = args.months.split(":")
    months = month_range(m_start, m_end)
    skip = set(args.skip.split(",")) if args.skip else set()

    data = json.load(open(args.companies, encoding="utf-8"))
    skipped = {"employment": 0, "financial": 0, "revenue": 0}

    def eok(v):
        return None if v is None else v / 1e8

    for c in data:
        name = c[args.name_field]

        if "employment" not in skip:
            p = make_employment_timeseries_chart(name, c, months, args.out_dir)
            c["__chart_employment"] = p
            if p is None:
                skipped["employment"] += 1

            # 고용 요약 통계 (표에 숫자로도 표시할 때 사용 — 17개월 전체를 표로 넣기엔 컬럼이 너무 많음)
            hc_pts, hires_sum, leaves_sum = [], 0, 0
            for m in months:
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

        if "financial" not in skip:
            fin_years = [(y,
                          eok(to_float(c.get(f"{y}년 자산"))),
                          eok(to_float(c.get(f"{y}년 부채"))),
                          eok(to_float(c.get(f"{y}년 자본")))) for y in years]
            p = make_grouped_bar_chart(name, "financial", fin_years, ["자산", "부채", "자본"], [NAVY, CORAL, TEAL], args.out_dir)
            c["__chart_financial"] = p
            if p is None:
                skipped["financial"] += 1

        if "revenue" not in skip:
            profit_years = [(y,
                             eok(to_float(c.get(f"{y}년 매출(원)"))),
                             eok(to_float(c.get(f"{y} 영업이익"))),
                             eok(to_float(c.get(f"{y} 순이익")))) for y in years]
            p = make_grouped_bar_chart(name, "profit", profit_years, ["매출액", "영업이익", "순이익"],
                                        [REV_TEAL, OPINC_BLUE, NETINC_GOLD], args.out_dir)
            c["__chart_revenue"] = p
            if p is None:
                skipped["revenue"] += 1

    json.dump(data, open(args.companies, "w", encoding="utf-8"), ensure_ascii=False, indent=2, default=str)
    print(f"완료: {len(data)}개 기업 처리")
    print(f"차트 스킵(데이터 2개년/개월 미만): {skipped}")


if __name__ == "__main__":
    main()
