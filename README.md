# 112384

ISO 학습 자료와 업무 자동화 도구 저장소. Claude 세션에 연결해 쓰도록 설정되어 있다.

## 구성
| 경로 | 역할 |
|---|---|
| `CLAUDE.md` | Claude가 항상 따르는 기본 지침 |
| `tools/asset_ledger.py` | 비품 대장 엑셀 생성·정리 (드롭다운, 처리 구분 색상, 자동 요약) |
| `tools/quote_compare.py` | 비교견적 계산 (공통 항목 총액, 렌탈 vs 구매, 대행사 vs 독립, 종합 점수, 차트) |
| `.claude/skills/asset-ledger/` | 비품 대장 작업 절차 |
| `.claude/skills/quote-compare/` | 비교견적서 작업 절차 |
| `.claude/skills/iso-model-answer/` | ISO 워크숍 모범답안 작성 절차 |
| `.claude/skills/clause-summary/` | ISO 조항별 요약 정리 절차 |
| `answers/`, `notes/` | ISO 답안·조항 노트 |

## 도구 사용법
```bash
pip install -r requirements.txt

# 비품 대장
python tools/asset_ledger.py --out 비품대장.xlsx                       # 빈 양식
python tools/asset_ledger.py --from 기존비품.xlsx --out 비품대장_정리.xlsx  # 기존 파일 정리

# 비교견적
python tools/quote_compare.py --template 견적입력.xlsx
python tools/quote_compare.py --in 견적입력.xlsx --out 비교견적.xlsx --png 비교차트.png
```

## 주의
이 저장소는 공개 상태다. 업체명·금액·직원 정보 같은 실제 업무 데이터는 올리지 않는다.
도구와 규칙만 저장소에 두고, 실제 파일은 대화에 첨부해서 처리한다.

## 고도화하는 법
- 지침을 바꾸려면 `CLAUDE.md`를 수정한다.
- 반복 작업이 생기면 `.claude/skills/<이름>/SKILL.md`를 추가한다.
- 계산·서식처럼 매번 똑같이 나와야 하는 작업은 `tools/`에 스크립트로 만든다.
