# 🛫 Aviation Dashboard — 국제선 항공통계

마이리얼트립 T&A 사업실 — 한국발 국제선 여객·운항·노선 통계를 팀별/도시별/월별로 시각화한 대시보드입니다.

## 🔗 라이브 대시보드

- **메인 (전체 팀)**: https://hyeonukjang-ui.github.io/aviation-dashboard/
- **중화권 전용**: https://hyeonukjang-ui.github.io/aviation-dashboard/china.html

## 📊 데이터 소스

[항공포털 — 항공통계상세조회](https://www.airportal.go.kr/stats/transport/chartDetail3.do)에서 김포(GMP)·인천(ICN) 노선별 월별 엑셀.

## 🔄 매월 데이터 업데이트하는 법

### 1. 항공포털에서 엑셀 다운로드

[항공포털](https://www.airportal.go.kr/stats/transport/chartDetail3.do)에서 매월:
- 김포공항 노선별 → 새 달 다운로드
- 인천공항 노선별 → 새 달 다운로드

### 2. data/raw/ 폴더에 추가

```
data/raw/
├── 김포 25년/     ← 2025년 김포 데이터 (1)~(12).xlsx
├── 김포 26년/     ← 2026년 김포 데이터 (1)~(n).xlsx ← 여기에 새 달 추가
├── 인천 25년/
└── 인천 26년/     ← 여기에도 새 달 추가
```

**파일명 규칙**: `항공통계상세조회(노선) (n).xlsx` — `(n)`이 월(1~12). 항공포털 다운로드 시 자동으로 이 패턴.

### 3. 스크립트 실행

```bash
python3 scripts/update_data.py
```

자동으로:
- ✅ data/raw/ 전체 엑셀 파싱
- ✅ 팀 매핑 (T&A 7개 팀)
- ✅ 다공항 도시 통합 (상하이=PVG+SHA, 도쿄=HND+NRT 등)
- ✅ YoY 계산 (직전 연도 동기간 vs 최신 연도)
- ✅ `docs/index.html` + `docs/china.html` 자동 갱신

### 4. GitHub Push

```bash
git add .
git commit -m "data: 26년 X월 추가"
git push
```

→ 1~2분 후 라이브 대시보드 자동 업데이트

## 📁 디렉토리 구조

```
aviation-dashboard/
├── docs/                      # GitHub Pages 루트
│   ├── index.html             # 메인 대시보드 (7개 팀 통합)
│   └── china.html             # 중화권 전용 대시보드
├── scripts/
│   └── update_data.py         # 매월 데이터 업데이트 스크립트
├── data/
│   ├── raw/                   # 원본 엑셀 (항공포털 다운로드)
│   │   ├── 김포 25년/
│   │   ├── 김포 26년/
│   │   ├── 인천 25년/
│   │   └── 인천 26년/
│   └── processed/             # 가공된 JSON (스크립트 자동 생성)
│       ├── main.json
│       └── china.json
└── README.md
```

## 🛠️ 의존성

```bash
pip3 install openpyxl
```

(또는 `python3 -m pip install openpyxl`)

## 📊 분석 항목

### 메인 대시보드
- 전체 KPI (총 여객 / YoY / 최성수기·비수기 / 최대 권역)
- 월별 추이 (GMP·ICN × 2025·2026)
- 7개 팀별 점유율 + YoY 막대
- **팀별 섹션 7개** (각 팀: 월별 추이 + 도시별 막대 + 도시×월 히트맵)
- 도시×월 히트맵에 **25년 / 26년 / YoY%** 동시 표시 (1~4월)
- 전체 도시 테이블 (정렬·검색)

### 중화권 전용 대시보드
- 한-중 무비자 효과 강조
- 5개 권역 분류 (중국 본토·홍콩·마카오·대만·싱가포르)
- 본토 49개 도시 성장 Top 20
- 권역별 시즈널리티

## 🎨 디자인

[마이리얼트립 시드 디자인 시스템](https://myrealtrip.github.io/tna-csm-playground/PDP_TF/Hyeonuk/seed-design.html) 적용:

| 토큰 | hex |
|---|---|
| Primary | `#51ABF3` (하늘색) |
| Success | `#33B893` |
| Warning | `#FFBF00` |
| Error | `#FA5B4A` |
| Text | `#212529` |

## 📝 라이센스

내부 사용 (T&A 사업실)
