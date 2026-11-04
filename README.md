# 나라장터 AI 사업 모니터

조달청 나라장터 OpenAPI로 **AI·RPA·빅데이터 등 키워드 관련 입찰공고와 낙찰결과를 매일 자동 수집**하고, 대시보드로 보여주는 도구입니다.

- 서버 없음: GitHub Actions가 매일 수집 → 저장소에 저장 → GitHub Pages가 화면 제공
- 외부 LLM 없음: 키워드·제외어 규칙으로만 필터링
- 공개 데이터만 다룸: 나라장터에 이미 공개된 공고·낙찰 정보

## 화면 구성
- 요약: 공고 건수 / 사업 규모 합계 / 낙찰 완료(낙찰총액·평균 낙찰률) / 최근 7일 신규
- 차트: 월별 공고 건수, 지역별 건수(누르면 지역 필터), 발주기관 TOP 10(누르면 기관 검색)
- 목록: 상태(진행중·결과대기·낙찰), 원문 링크, 낙찰업체·금액·낙찰률, 정렬, 엑셀 다운로드
- 데이터가 아직 없으면 가상의 샘플 데이터로 화면을 미리 보여줍니다.

## 설치 순서 (약 20분)

### 1) API 활용신청 — 공공데이터포털
아래 두 서비스를 각각 **활용신청**합니다. (일반 인증키 하나로 둘 다 사용)
- `조달청_나라장터 입찰공고정보서비스`
- `조달청_나라장터 낙찰정보서비스`

마이페이지 › 개발계정에서 **일반 인증키(Decoding)** 를 복사해 둡니다.

### 2) GitHub 저장소 만들기
1. 새 저장소 생성 (Public 권장 — 무료 계정은 Public 저장소만 Pages 사용 가능. 담긴 데이터는 모두 공개 정보입니다)
2. 이 폴더의 파일 전체를 업로드 (`.github` 폴더 포함)

### 3) API 키 등록 — Secrets
저장소 › Settings › Secrets and variables › Actions › **New repository secret**
- Name: `G2B_SERVICE_KEY`
- Secret: 복사한 인증키

> 키는 Secrets에만 넣고, 코드나 config.json에는 절대 쓰지 마세요.

### 4) 화면 켜기 — Pages
Settings › Pages › Source: **Deploy from a branch** › Branch: `main` / `(root)` › Save
→ 1~2분 뒤 `https://<아이디>.github.io/<저장소명>/` 에서 대시보드가 열립니다.

### 5) 첫 수집
Actions 탭 › **나라장터 공고 수집** › Run workflow › backfill 칸에 `90` 입력 › 실행
→ 최근 90일 공고가 쌓입니다. 이후에는 매일 오전 6시 47분에 자동으로 최근 3일치를 갱신합니다.

## 설정 바꾸기 (`config.json`)
| 항목 | 의미 |
|---|---|
| `keywords` | 수집할 공고명 키워드. 바꾸면 다른 주제 모니터링기가 됩니다 |
| `exclude_keywords` | 공고명에 있으면 제외할 단어 |
| `bid_types` | `용역`, `물품`, `공사` 중 선택 |
| `daily_lookback_days` | 매일 다시 훑는 기간(정정공고·낙찰 반영용) |
| `backfill_days` | 데이터가 비었을 때 처음 가져오는 기간 |

영문 약어 키워드(AI, RPA 등)는 앞뒤가 영문자일 때(MAIL, TRAIN, AIR 등) 자동으로 걸러집니다.

## 로컬에서 시험해 보기
```bash
pip install requests
export G2B_SERVICE_KEY="인증키"
python collector/collect.py --backfill 7 --dry-run   # 저장 없이 결과 확인
python collector/collect.py --backfill 30            # data/bids.json 저장
python -m http.server 8000                           # http://localhost:8000 에서 화면 확인
```

## 알아둘 점
- **지역은 추정값**입니다. 수요기관명에 시도·주요 시 이름이 있으면 그 지역, 없으면 `중앙·기타`.
- **낙찰 연결**은 공고번호 기준입니다. 낙찰 API를 같은 키워드로 조회해 이미 수집된 공고에 붙입니다.
- API는 한 번에 1개월 이내만 조회되므로 수집기가 28일 단위로 나눠 호출합니다.
- 인증키 오류 등으로 **모든 호출이 실패하면 기존 데이터를 덮어쓰지 않고** 작업이 실패로 끝납니다. (Actions 탭에서 확인)
- 활용신청 직후에는 키가 등록되기까지 1시간 정도 걸릴 수 있습니다 (`SERVICE_KEY_IS_NOT_REGISTERED_ERROR`).

## 파일 구조
```
├─ index.html                 대시보드 (GitHub Pages)
├─ config.json                키워드·수집 설정
├─ collector/collect.py       수집기
├─ data/bids.json             수집 결과 (자동 갱신)
└─ .github/workflows/collect.yml   매일 자동 실행
```
