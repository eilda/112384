"""비교견적서 자동 계산 도구.

사용법
  1) 입력 양식 만들기:  python tools/quote_compare.py --template 견적입력.xlsx
  2) 제안서 내용을 양식에 한 줄 = 한 업체의 한 항목으로 입력
  3) 비교 실행:         python tools/quote_compare.py --in 견적입력.xlsx --out 비교견적.xlsx --png 비교차트.png
     가중치 변경:       --weights 가격=50,납기=20,포함범위=15,AS=15
     렌탈 판단 기준:    --rental-ok 10 --rental-max 20  (렌탈이 구매보다 n% 이내로 비싸면 렌탈)

계산 내용
  - 항목별 구매 금액(수량×단가)과 렌탈 금액(수량×월액×개월)
  - 렌탈 vs 구매 판정: 차이 ≤ rental-ok% → 렌탈 권장, ≤ rental-max% → 렌탈 검토, 초과 → 구매 권장
  - 업체별 총액, 공통 항목만으로 비교한 총액(누락 항목 때문에 싸 보이는 착시 방지), 누락 항목 목록
  - 항목별 최저가 표시, 최저가 대비 차이(%)
  - 대행사 견적이 독립 견적 최저가보다 얼마나 높은지
  - 가중치 종합 점수와 추천 업체
"""
import argparse
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

IN_COLS = [
    "업체", "견적 출처", "항목", "수량", "구매 단가", "렌탈 월액", "렌탈 개월",
    "설치 포함", "철거 포함", "AS 기간(개월)", "납기 가능", "비고",
]
GUIDE = [
    ("업체", "업체 이름 (같은 업체는 같은 이름으로)"),
    ("견적 출처", "대행사 / 독립 중 선택"),
    ("항목", "업체끼리 같은 품목은 같은 이름으로 입력해야 비교됩니다 (예: 회의 테이블 1800)"),
    ("수량", "숫자"),
    ("구매 단가", "1개당 구매 가격(원). 렌탈만 제안했으면 비워 둠"),
    ("렌탈 월액", "1개당 월 렌탈료(원). 구매만 제안했으면 비워 둠"),
    ("렌탈 개월", "렌탈 계약 기간(개월)"),
    ("설치 포함 / 철거 포함", "Y / N"),
    ("AS 기간(개월)", "무상 AS 기간"),
    ("납기 가능", "이전 일정에 맞출 수 있으면 Y"),
]
NAVY = "1F3A5F"
GREEN = "DDF3E4"
AMBER = "FFF1D6"
RED = "FBE0E0"
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
HEADER_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="C9CED6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ---------- 입력 ----------
def make_template(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "견적입력"
    ws.append(IN_COLS)
    widths = [14, 10, 24, 7, 12, 12, 10, 9, 9, 11, 9, 24]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        c = ws.cell(row=1, column=i)
        c.fill, c.font, c.border = HEADER_FILL, HEADER_FONT, BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for col, opts in (("B", "대행사,독립"), ("H", "Y,N"), ("I", "Y,N"), ("K", "Y,N")):
        dv = DataValidation(type="list", formula1=f'"{opts}"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}500")
    for col in "EF":
        for r in range(2, 501):
            ws[f"{col}{r}"].number_format = "#,##0"
    ws.freeze_panes = "A2"
    g = wb.create_sheet("작성 안내")
    g.append(["열", "입력 방법"])
    for row in GUIDE:
        g.append(list(row))
    g.column_dimensions["A"].width = 22
    g.column_dimensions["B"].width = 80
    for c in g[1]:
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    wb.save(path)


def yn(v):
    return str(v).strip().upper() in ("Y", "O", "YES", "예", "포함", "가능", "TRUE", "1")


def num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, str):
        v = v.replace(",", "").replace("원", "").strip()
        if not v:
            return None
    try:
        return float(v)
    except ValueError:
        return None


def load(path):
    df = pd.read_excel(path, sheet_name=0, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in ("업체", "항목") if c not in df.columns]
    if missing:
        sys.exit("입력 파일에 필요한 열이 없습니다: " + ", ".join(missing))
    for c in IN_COLS:
        if c not in df.columns:
            df[c] = None
    df = df[df["업체"].notna() & df["항목"].notna()].copy()
    df["업체"] = df["업체"].astype(str).str.strip()
    df["항목"] = df["항목"].astype(str).str.strip()
    df["견적 출처"] = df["견적 출처"].fillna("독립").astype(str).str.strip()
    df["수량"] = df["수량"].map(num).fillna(1)
    for c in ("구매 단가", "렌탈 월액", "렌탈 개월", "AS 기간(개월)"):
        df[c] = df[c].map(num)
    for c in ("설치 포함", "철거 포함", "납기 가능"):
        df[c] = df[c].map(yn)
    return df.reset_index(drop=True)


# ---------- 계산 ----------
def judge(buy, rent, ok, mx):
    if buy is None and rent is None:
        return None, None, "가격 없음"
    if rent is None:
        return buy, None, "구매만 제안"
    if buy is None:
        return rent, None, "렌탈만 제안"
    diff = (rent - buy) / buy * 100
    if diff <= ok:
        return rent, diff, "렌탈 권장"
    if diff <= mx:
        return rent, diff, "렌탈 검토"
    return buy, diff, "구매 권장"


def compute(df, ok, mx, weights):
    df["구매 금액"] = df.apply(lambda r: r["수량"] * r["구매 단가"] if r["구매 단가"] is not None else None, axis=1)
    df["렌탈 금액"] = df.apply(
        lambda r: r["수량"] * r["렌탈 월액"] * (r["렌탈 개월"] or 0)
        if r["렌탈 월액"] is not None and r["렌탈 개월"] else None, axis=1)
    res = df.apply(lambda r: judge(r["구매 금액"], r["렌탈 금액"], ok, mx), axis=1)
    df["비교 금액"] = [x[0] for x in res]
    df["렌탈-구매 차이(%)"] = [x[1] for x in res]
    df["판정"] = [x[2] for x in res]

    vendors = list(dict.fromkeys(df["업체"]))
    items = list(dict.fromkeys(df["항목"]))
    pivot = df.pivot_table(index="항목", columns="업체", values="비교 금액", aggfunc="sum").reindex(items)
    pivot = pivot.reindex(columns=vendors)
    common = [i for i in items if pivot.loc[i].notna().all()]

    rows = []
    for v in vendors:
        d = df[df["업체"] == v]
        quoted = set(d["항목"])
        rows.append({
            "업체": v,
            "견적 출처": d["견적 출처"].iloc[0],
            "제안 항목 수": len(quoted),
            "누락 항목": ", ".join(i for i in items if i not in quoted) or "-",
            "전체 총액": d["비교 금액"].sum(skipna=True),
            "공통 항목 총액": d[d["항목"].isin(common)]["비교 금액"].sum(skipna=True),
            "납기 충족(%)": d["납기 가능"].mean() * 100,
            "설치·철거 포함(%)": (d["설치 포함"].mean() + d["철거 포함"].mean()) / 2 * 100,
            "평균 AS(개월)": d["AS 기간(개월)"].mean() if d["AS 기간(개월)"].notna().any() else 0,
        })
    summary = pd.DataFrame(rows)

    base = "공통 항목 총액" if common else "전체 총액"
    lowest = summary[base][summary[base] > 0].min()
    max_as = summary["평균 AS(개월)"].max() or 1
    summary["가격 점수"] = summary[base].map(lambda x: lowest / x * 100 if x else 0)
    summary["납기 점수"] = summary["납기 충족(%)"]
    summary["포함범위 점수"] = summary["설치·철거 포함(%)"]
    summary["AS 점수"] = summary["평균 AS(개월)"] / max_as * 100
    wsum = sum(weights.values())
    summary["종합 점수"] = (
        summary["가격 점수"] * weights["가격"] + summary["납기 점수"] * weights["납기"]
        + summary["포함범위 점수"] * weights["포함범위"] + summary["AS 점수"] * weights["AS"]
    ) / wsum
    summary = summary.sort_values("종합 점수", ascending=False).reset_index(drop=True)
    summary.insert(0, "순위", range(1, len(summary) + 1))

    # 대행사 vs 독립 최저가
    agency = df[df["견적 출처"] == "대행사"]
    indep = df[df["견적 출처"] != "대행사"]
    gap = []
    if not agency.empty and not indep.empty:
        indep_min = indep.groupby("항목")["비교 금액"].min()
        for _, r in agency.iterrows():
            m = indep_min.get(r["항목"])
            if m and r["비교 금액"]:
                gap.append({"항목": r["항목"], "대행사 업체": r["업체"], "대행사 금액": r["비교 금액"],
                            "독립 최저가": m, "차이(%)": (r["비교 금액"] - m) / m * 100})
    gap = pd.DataFrame(gap)
    return df, summary, pivot, common, base, gap


# ---------- 출력 ----------
def header(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font, cell.border = HEADER_FILL, HEADER_FONT, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_df(ws, df, start_row=1, money=(), pct=(), one=()):
    for c, name in enumerate(df.columns, 1):
        ws.cell(row=start_row, column=c, value=name)
    header(ws, start_row, len(df.columns))
    for i, rec in enumerate(df.itertuples(index=False), start_row + 1):
        for c, (name, val) in enumerate(zip(df.columns, rec), 1):
            if isinstance(val, float) and pd.isna(val):
                val = None
            cell = ws.cell(row=i, column=c, value=val)
            cell.border = BORDER
            if name in money:
                cell.number_format = "#,##0"
            elif name in pct:
                cell.number_format = "0.0"
            elif name in one:
                cell.number_format = "0.0"
    return start_row + len(df)


def autosize(ws, minimum=9, maximum=40):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        longest = max((len(str(c.value)) * (1.8 if any(ord(ch) > 127 for ch in str(c.value)) else 1.1)
                       for c in col if c.value is not None), default=minimum)
        ws.column_dimensions[letter].width = max(minimum, min(maximum, longest + 2))


def fmt_won(x):
    return f"{x:,.0f}원"


def build_report(df, summary, pivot, common, base, gap, weights, ok, mx, out):
    wb = Workbook()
    s = wb.active
    s.title = "요약"
    top = summary.iloc[0]
    s["A1"] = "비교견적 요약"
    s["A1"].font = Font(bold=True, size=16, color=NAVY)
    s["A3"] = f"추천: {top['업체']}"
    s["A3"].font = Font(bold=True, size=13, color="1E7B3C")
    lines = [
        f"종합 점수 {top['종합 점수']:.1f}점 (가중치 가격 {weights['가격']} · 납기 {weights['납기']} · "
        f"포함범위 {weights['포함범위']} · AS {weights['AS']})",
        f"{base} {fmt_won(top[base])}",
    ]
    if len(summary) > 1:
        second = summary.iloc[1]
        d = second[base] - top[base]
        lines.append(f"2순위 {second['업체']} 대비 {fmt_won(abs(d))} {'저렴' if d > 0 else '비쌈'}")
    if common and len(common) < df["항목"].nunique():
        lines.append(f"공통 항목 {len(common)}개 기준으로 가격 비교 (일부 업체 누락 항목 있음 → 업체별 비교 시트 확인)")
    if not gap.empty:
        lines.append(f"대행사 견적은 독립 견적 최저가보다 평균 {gap['차이(%)'].mean():.1f}% 높음")
    cnt = df["판정"].value_counts()
    if cnt.get("렌탈 권장", 0) or cnt.get("렌탈 검토", 0):
        lines.append(f"렌탈 권장 {cnt.get('렌탈 권장', 0)}건 · 렌탈 검토 {cnt.get('렌탈 검토', 0)}건 "
                     f"(렌탈이 구매보다 {ok}% 이내 → 권장, {mx}% 이내 → 검토)")
    for i, t in enumerate(lines, 4):
        s.cell(row=i, column=1, value="• " + t)

    r0 = len(lines) + 6
    cols = ["순위", "업체", "견적 출처", base, "종합 점수", "가격 점수", "납기 점수", "포함범위 점수", "AS 점수"]
    last = write_df(s, summary[cols], r0, money=(base,), one=cols[4:])
    for c in range(1, len(cols) + 1):
        s.cell(row=r0 + 1, column=c).fill = PatternFill("solid", fgColor=GREEN)
        s.cell(row=r0 + 1, column=c).font = Font(bold=True)

    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = f"업체별 {base}"
    chart.y_axis.title = "원"
    chart.y_axis.numFmt = "#,##0"
    chart.legend = None
    data = Reference(s, min_col=4, min_row=r0, max_row=last)
    cats = Reference(s, min_col=2, min_row=r0 + 1, max_row=last)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height, chart.width = 7, 16
    s.add_chart(chart, f"A{last + 3}")
    autosize(s)
    s.column_dimensions["A"].width = 8

    v = wb.create_sheet("업체별 비교")
    vcols = ["순위", "업체", "견적 출처", "제안 항목 수", "누락 항목", "전체 총액", "공통 항목 총액",
             "납기 충족(%)", "설치·철거 포함(%)", "평균 AS(개월)", "종합 점수"]
    write_df(v, summary[vcols], 1, money=("전체 총액", "공통 항목 총액"),
             one=("납기 충족(%)", "설치·철거 포함(%)", "평균 AS(개월)", "종합 점수"))
    for i in range(2, len(summary) + 2):
        if v.cell(row=i, column=5).value not in ("-", None):
            v.cell(row=i, column=5).fill = PatternFill("solid", fgColor=AMBER)
    autosize(v)

    p = wb.create_sheet("항목별 비교")
    pv = pivot.copy()
    pv["최저가"] = pv.min(axis=1)
    pv["최고가 대비 절감액"] = pv.drop(columns="최저가").max(axis=1) - pv["최저가"]
    pv = pv.reset_index()
    vendors = [c for c in pivot.columns]
    write_df(p, pv, 1, money=tuple(vendors) + ("최저가", "최고가 대비 절감액"))
    for i in range(2, len(pv) + 2):
        vals = {c: p.cell(row=i, column=j + 2).value for j, c in enumerate(vendors)}
        present = [x for x in vals.values() if x is not None]
        lo = min(present) if present else None
        for j, c in enumerate(vendors):
            cell = p.cell(row=i, column=j + 2)
            if cell.value is None:
                cell.value = "누락"
                cell.fill = PatternFill("solid", fgColor=AMBER)
                cell.alignment = Alignment(horizontal="center")
            elif cell.value == lo:
                cell.fill = PatternFill("solid", fgColor=GREEN)
                cell.font = Font(bold=True)
    autosize(p)

    rb = wb.create_sheet("렌탈 vs 구매")
    rcols = ["업체", "항목", "수량", "구매 금액", "렌탈 금액", "렌탈-구매 차이(%)", "판정"]
    write_df(rb, df[rcols], 1, money=("구매 금액", "렌탈 금액"), one=("렌탈-구매 차이(%)",))
    colors = {"렌탈 권장": GREEN, "렌탈 검토": AMBER, "구매 권장": RED}
    for i in range(2, len(df) + 2):
        cell = rb.cell(row=i, column=7)
        if cell.value in colors:
            cell.fill = PatternFill("solid", fgColor=colors[cell.value])
    autosize(rb)

    if not gap.empty:
        g = wb.create_sheet("대행사 vs 독립")
        write_df(g, gap, 1, money=("대행사 금액", "독립 최저가"), one=("차이(%)",))
        for i in range(2, len(gap) + 2):
            cell = g.cell(row=i, column=5)
            if cell.value and cell.value > 10:
                cell.fill = PatternFill("solid", fgColor=RED)
        autosize(g)

    raw = wb.create_sheet("입력 원본")
    write_df(raw, df[IN_COLS + ["구매 금액", "렌탈 금액", "비교 금액"]], 1,
             money=("구매 단가", "렌탈 월액", "구매 금액", "렌탈 금액", "비교 금액"))
    autosize(raw)
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = ws.title != "요약"
    wb.save(out)


def build_png(summary, base, path):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager

    for name in ("Noto Sans CJK KR", "NanumGothic", "Malgun Gothic", "AppleGothic"):
        if any(name in f.name for f in font_manager.fontManager.ttflist):
            plt.rcParams["font.family"] = name
            break
    plt.rcParams["axes.unicode_minus"] = False

    d = summary.iloc[::-1]
    colors = ["#1E7B3C" if r == 1 else ("#B7791F" if s == "대행사" else "#8A94A6")
              for r, s in zip(d["순위"], d["견적 출처"])]
    fig, axes = plt.subplots(1, 2, figsize=(11, max(3, 0.7 * len(d) + 1.5)))
    ax = axes[0]
    bars = ax.barh(d["업체"], d[base], color=colors)
    ax.set_title(base, fontsize=13, fontweight="bold")
    top = d[base].max() or 1
    for b, val in zip(bars, d[base]):
        ax.text(b.get_width() + top * 0.01, b.get_y() + b.get_height() / 2, f"{val / 10000:,.0f}만원",
                va="center", fontsize=10)
    ax.set_xlim(0, top * 1.22)
    ax.xaxis.set_visible(False)

    ax = axes[1]
    bars = ax.barh(d["업체"], d["종합 점수"], color=colors)
    ax.set_title("종합 점수", fontsize=13, fontweight="bold")
    for b, val in zip(bars, d["종합 점수"]):
        ax.text(b.get_width() + 1, b.get_y() + b.get_height() / 2, f"{val:.1f}", va="center", fontsize=10)
    ax.set_xlim(0, 115)
    ax.xaxis.set_visible(False)
    for a in axes:
        for side in ("top", "right", "bottom"):
            a.spines[side].set_visible(False)
    fig.suptitle(f"추천: {summary.iloc[0]['업체']}", fontsize=15, fontweight="bold", color="#1E7B3C", x=0.02,
                 ha="left")
    fig.text(0.02, 0.01, "초록 = 추천 · 갈색 = 대행사 견적 · 회색 = 기타", fontsize=9, color="#555")
    fig.tight_layout(rect=(0, 0.04, 1, 0.92))
    fig.savefig(path, dpi=200)


def parse_weights(text):
    w = {"가격": 50, "납기": 20, "포함범위": 15, "AS": 15}
    if text:
        for part in text.split(","):
            k, v = part.split("=")
            k = k.strip()
            if k not in w:
                sys.exit(f"가중치 이름은 {', '.join(w)} 중 하나여야 합니다: {k}")
            w[k] = float(v)
    return w


def main():
    p = argparse.ArgumentParser(description="비교견적서 자동 계산")
    p.add_argument("--template", help="빈 입력 양식을 만들 파일 이름")
    p.add_argument("--in", dest="src", help="견적 입력 파일")
    p.add_argument("--out", help="결과 엑셀 파일 이름")
    p.add_argument("--png", help="보고용 차트 이미지 파일 이름 (선택)")
    p.add_argument("--weights", help="예: 가격=50,납기=20,포함범위=15,AS=15")
    p.add_argument("--rental-ok", type=float, default=10, help="렌탈 권장 기준 % (기본 10)")
    p.add_argument("--rental-max", type=float, default=20, help="렌탈 검토 상한 % (기본 20)")
    a = p.parse_args()

    if a.template:
        make_template(a.template)
        print(f"입력 양식 생성: {a.template}")
        return
    if not (a.src and a.out):
        p.error("--template 또는 --in 과 --out 을 지정하세요")
    weights = parse_weights(a.weights)
    df = load(a.src)
    if df.empty:
        sys.exit("입력된 견적이 없습니다.")
    df, summary, pivot, common, base, gap = compute(df, a.rental_ok, a.rental_max, weights)
    build_report(df, summary, pivot, common, base, gap, weights, a.rental_ok, a.rental_max, a.out)
    print(f"완료: {a.out}")
    if a.png:
        build_png(summary, base, a.png)
        print(f"차트: {a.png}")
    top = summary.iloc[0]
    print(f"추천: {top['업체']} (종합 {top['종합 점수']:.1f}점, {base} {top[base]:,.0f}원)")


if __name__ == "__main__":
    sys.exit(main())
