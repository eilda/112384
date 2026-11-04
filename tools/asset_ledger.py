"""비품 대장 엑셀 생성·고도화 도구.

사용법
  빈 양식 만들기:      python tools/asset_ledger.py --out 비품대장.xlsx
  기존 파일 정리하기:  python tools/asset_ledger.py --from 기존비품.xlsx --out 비품대장_정리.xlsx
                       (--sheet 시트이름 으로 시트 지정 가능)

기능
  - 표준 열 구조로 정리 (기존 열 이름은 비슷한 이름끼리 자동 매칭, 매칭 안 된 열은 뒤에 보존)
  - 드롭다운: 분류 / 상태 / 처리 구분(이전·매각·폐기·보류)
  - 처리 구분별 행 색상, 고장·분실은 빨간 글씨
  - 관리번호 빈칸 자동 채번 (A-0001 …)
  - 요약 시트: 처리 구분별 건수·수량·구입가 합계, 분류×처리 구분 표, 처리 구분 미입력 건수 (수식이라 자동 갱신)
"""
import argparse
import re
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

COLUMNS = [
    "관리번호", "분류", "품목", "모델/규격", "수량", "상태", "현 위치",
    "사용자/담당", "처리 구분", "신사옥 배치", "구입일", "구입가(원)", "비고",
]
WIDTHS = {
    "관리번호": 11, "분류": 10, "품목": 22, "모델/규격": 22, "수량": 7, "상태": 10,
    "현 위치": 14, "사용자/담당": 13, "처리 구분": 11, "신사옥 배치": 14,
    "구입일": 12, "구입가(원)": 13, "비고": 28,
}
CHOICES = {
    "분류": ["IT기기", "가구", "가전", "사무용품", "기타"],
    "상태": ["정상", "수리필요", "고장", "분실"],
    "처리 구분": ["이전", "매각", "폐기", "보류"],
}
SYNONYMS = {
    "관리번호": ["관리번호", "자산번호", "비품번호", "일련번호", "번호", "no", "id"],
    "분류": ["분류", "구분", "카테고리", "종류", "유형", "대분류"],
    "품목": ["품목", "품명", "비품명", "자산명", "장비명", "제품명", "기기명", "물품명"],
    "모델/규격": ["모델", "모델명", "규격", "사양", "스펙", "시리얼", "시리얼번호", "s/n"],
    "수량": ["수량", "개수", "대수", "갯수", "qty"],
    "상태": ["상태", "상태값", "사용상태"],
    "현 위치": ["위치", "현위치", "장소", "보관위치", "설치장소", "층"],
    "사용자/담당": ["사용자", "담당자", "담당", "사용부서", "부서", "관리자"],
    "처리 구분": ["처리구분", "처리", "조치", "이전여부", "처분", "처리방법"],
    "신사옥 배치": ["신사옥", "신사옥배치", "배치", "이전위치"],
    "구입일": ["구입일", "취득일", "구매일", "입고일", "도입일"],
    "구입가(원)": ["구입가", "취득가", "구매가", "금액", "단가", "가격", "취득금액"],
    "비고": ["비고", "메모", "참고", "특이사항"],
}

ROW_FILL = {
    "이전": "DCEBFB", "매각": "FFF1D6", "폐기": "FBE0E0", "보류": "ECECEC",
}
HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
HEADER_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="C9CED6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def norm(text):
    return re.sub(r"[\s_\-()/\[\].·:]", "", str(text)).lower()


def match_column(header):
    h = norm(header)
    if not h:
        return None
    for target, words in SYNONYMS.items():
        if h in (norm(w) for w in words):
            return target
    for target, words in SYNONYMS.items():
        if any(len(norm(w)) >= 2 and norm(w) in h for w in words):
            return target
    if len(h) >= 2:  # '특이' → '특이사항'처럼 줄여 쓴 열 이름
        for target, words in SYNONYMS.items():
            if any(h in norm(w) for w in words):
                return target
    return None


def find_header_row(raw):
    """앞쪽 10행 중 표준 열과 가장 많이 맞는 행을 머리글로 본다."""
    best, best_hits = 0, -1
    for i in range(min(10, len(raw))):
        hits = sum(1 for v in raw.iloc[i].tolist() if match_column(v))
        if hits > best_hits:
            best, best_hits = i, hits
    return best


def load_existing(path, sheet):
    raw = pd.read_excel(path, sheet_name=sheet or 0, header=None, dtype=object)
    raw = raw.dropna(how="all").reset_index(drop=True)
    hdr = find_header_row(raw)
    headers = [str(v).strip() if pd.notna(v) else "" for v in raw.iloc[hdr].tolist()]
    body = raw.iloc[hdr + 1:].reset_index(drop=True)
    body.columns = range(len(headers))

    mapping, extras, used = {}, [], set()
    for idx, h in enumerate(headers):
        target = match_column(h)
        if target and target not in used:
            mapping[target] = idx
            used.add(target)
        elif h or body[idx].notna().any():
            extras.append((h or f"열{idx + 1}", idx))

    rows = []
    for _, r in body.iterrows():
        if r.isna().all():
            continue
        rec = {c: (r[mapping[c]] if c in mapping else None) for c in COLUMNS}
        for name, idx in extras:
            rec[name] = r[idx]
        rows.append(rec)
    report = {
        "header_row": hdr + 1,
        "mapped": {k: headers[v] for k, v in mapping.items()},
        "extras": [n for n, _ in extras],
    }
    return rows, [n for n, _ in extras], report


def clean(value):
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().date()
    return value


def build(rows, extras, out):
    wb = Workbook()
    ws = wb.active
    ws.title = "비품대장"
    cols = COLUMNS + extras
    ws.append(cols)
    for c, name in enumerate(cols, 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font, cell.border = HEADER_FILL, HEADER_FONT, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = WIDTHS.get(name, 16)
    ws.row_dimensions[1].height = 28

    seq = 0
    for rec in rows:
        values = [clean(rec.get(c)) for c in cols]
        if values[0] in (None, ""):
            seq += 1
            values[0] = f"A-{seq:04d}"
        ws.append(values)

    last = max(ws.max_row, 2)
    span = last + 300  # 행을 추가해도 서식·드롭다운이 적용되도록 여유
    col_of = {name: get_column_letter(i + 1) for i, name in enumerate(cols)}
    end_col = get_column_letter(len(cols))

    for name, options in CHOICES.items():
        dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
        dv.error, dv.errorTitle = "목록에서 선택해 주세요: " + ", ".join(options), "입력값 확인"
        ws.add_data_validation(dv)
        dv.add(f"{col_of[name]}2:{col_of[name]}{span}")

    proc = col_of["처리 구분"]
    for label, color in ROW_FILL.items():
        ws.conditional_formatting.add(
            f"A2:{end_col}{span}",
            FormulaRule(formula=[f'${proc}2="{label}"'], fill=PatternFill("solid", fgColor=color)),
        )
    st = col_of["상태"]
    ws.conditional_formatting.add(
        f"{st}2:{st}{span}",
        FormulaRule(formula=[f'OR(${st}2="고장",${st}2="분실")'], font=Font(color="C00000", bold=True)),
    )

    for row in ws.iter_rows(min_row=2, max_row=last, max_col=len(cols)):
        for cell in row:
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    for r in range(2, span + 1):
        ws[f"{col_of['구입가(원)']}{r}"].number_format = "#,##0"
        ws[f"{col_of['구입일']}{r}"].number_format = "yyyy-mm-dd"
        ws[f"{col_of['수량']}{r}"].alignment = Alignment(horizontal="center")

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{end_col}{last}"
    build_summary(wb, col_of, span)
    wb.save(out)
    return len(rows)


def build_summary(wb, col_of, span):
    s = wb.create_sheet("요약", 0)
    src = "비품대장"
    rng = lambda name: f"{src}!${col_of[name]}$2:${col_of[name]}${span}"  # noqa: E731
    title_font = Font(bold=True, size=14, color="1F3A5F")
    s["A1"], s["A1"].font = "비품 현황 요약 (자동 계산)", title_font
    s["A2"] = "비품대장 시트를 수정하면 자동으로 바뀝니다."
    s["A2"].font = Font(color="777777", size=9)

    s.append([])
    s.append(["처리 구분", "건수", "수량 합계", "구입가 합계(원)"])
    head_row = s.max_row
    for label in CHOICES["처리 구분"]:
        r = s.max_row + 1
        s.append([
            label,
            f'=COUNTIF({rng("처리 구분")},"{label}")',
            f'=SUMIF({rng("처리 구분")},"{label}",{rng("수량")})',
            f'=SUMIF({rng("처리 구분")},"{label}",{rng("구입가(원)")})',
        ])
        s[f"A{r}"].fill = PatternFill("solid", fgColor=ROW_FILL[label])
    r = s.max_row + 1
    s.append([
        "미입력", f'=COUNTIFS({rng("품목")},"<>",{rng("처리 구분")},"")',
        f'=SUMIFS({rng("수량")},{rng("품목")},"<>",{rng("처리 구분")},"")', None,
    ])
    s[f"A{r}"].font = Font(color="C00000", bold=True)
    total = s.max_row + 1
    s.append(["합계", f"=SUM(B{head_row + 1}:B{total - 1})", f"=SUM(C{head_row + 1}:C{total - 1})",
              f"=SUM(D{head_row + 1}:D{total - 1})"])
    style_table(s, head_row, total, 4)

    s.append([])
    s.append(["분류 \\ 처리 구분"] + CHOICES["처리 구분"] + ["미입력"])
    head2 = s.max_row
    for cat in CHOICES["분류"]:
        row = [cat]
        for label in CHOICES["처리 구분"]:
            row.append(f'=COUNTIFS({rng("분류")},"{cat}",{rng("처리 구분")},"{label}")')
        row.append(f'=COUNTIFS({rng("분류")},"{cat}",{rng("처리 구분")},"")')
        s.append(row)
    style_table(s, head2, s.max_row, len(CHOICES["처리 구분"]) + 2)

    s.append([])
    s.append(["상태 점검", "건수"])
    head3 = s.max_row
    for label in ["수리필요", "고장", "분실"]:
        s.append([label, f'=COUNTIF({rng("상태")},"{label}")'])
    style_table(s, head3, s.max_row, 2)

    for col, w in zip("ABCDEF", [18, 10, 12, 16, 10, 10]):
        s.column_dimensions[col].width = w


def style_table(ws, head, last, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=head, column=c)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for r in range(head, last + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            if r > head and c > 1:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right")
    for c in range(1, ncols + 1):
        if ws.cell(row=last, column=1).value == "합계":
            ws.cell(row=last, column=c).font = Font(bold=True)


def main():
    p = argparse.ArgumentParser(description="비품 대장 엑셀 생성·정리")
    p.add_argument("--from", dest="src", help="정리할 기존 엑셀 파일")
    p.add_argument("--sheet", help="기존 파일의 시트 이름 (기본: 첫 시트)")
    p.add_argument("--out", required=True, help="저장할 파일 이름")
    a = p.parse_args()

    if a.src:
        rows, extras, report = load_existing(a.src, a.sheet)
        print(f"머리글 행: {report['header_row']}행")
        for k, v in report["mapped"].items():
            print(f"  {v!r} → {k}")
        missing = [c for c in COLUMNS if c not in report["mapped"]]
        if missing:
            print("  새로 추가된 빈 열: " + ", ".join(missing))
        if report["extras"]:
            print("  매칭 안 돼 뒤에 보존한 열: " + ", ".join(report["extras"]))
    else:
        rows, extras = [], []
    n = build(rows, extras, a.out)
    print(f"완료: {a.out} (비품 {n}건)")


if __name__ == "__main__":
    sys.exit(main())
