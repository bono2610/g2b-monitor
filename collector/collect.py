#!/usr/bin/env python3
"""
나라장터 공고 모니터링 수집기

- 조달청 나라장터 OpenAPI(공공데이터포털)에서 키워드별 입찰공고와 낙찰결과를 수집
- 공고번호로 둘을 연결해 data/bids.json 에 누적 저장
- 외부 LLM 없이 규칙(키워드/제외어)만으로 필터링
- 지역·기관구분·분야 분류와 재공고 중복 정리도 config.json 규칙으로 수집 단계에서 계산

사용법
  export G2B_SERVICE_KEY="공공데이터포털 일반 인증키(Decoding)"
  python collector/collect.py              # 평소: 최근 N일(daily_lookback_days) 수집
  python collector/collect.py --backfill 90  # 최초: 최근 90일 몰아서 수집
  python collector/collect.py --dry-run    # 저장하지 않고 결과만 출력
  python collector/collect.py --reprocess  # API 호출 없이 기존 데이터에 분류·제외·중복 규칙만 다시 적용
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import unquote

import requests

from rules import OTHER_REGION, Rules, postprocess, rules_meta

ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT / "config.json"
DATA_PATH = ROOT / "data" / "bids.json"
KST = timezone(timedelta(hours=9))

NUM_ROWS = 100
MAX_PAGES = 30          # 키워드·기간당 최대 3,000건
CHUNK_DAYS = 28         # API 조회기간은 보통 1개월 이내만 허용
REQUEST_GAP = 0.25      # 호출 간격(초) — 트래픽 제한 대비

# ───────────────────────── 유틸 ─────────────────────────
def load_config() -> dict:
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def get_service_key() -> str:
    key = os.environ.get("G2B_SERVICE_KEY", "").strip()
    if not key:
        sys.exit("❌ 환경변수 G2B_SERVICE_KEY 가 비어 있습니다. (GitHub Secrets 또는 로컬 export 확인)")
    # Encoding 키(%2B 등 포함)를 넣어도 동작하도록 원본으로 되돌림
    return unquote(key) if "%" in key else key


def to_int(v) -> int | None:
    if v is None:
        return None
    s = str(v).replace(",", "").strip()
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def to_float(v) -> float | None:
    s = str(v or "").replace(",", "").replace("%", "").strip()
    try:
        return float(s) if s else None
    except ValueError:
        return None


def fmt_dt(dt: datetime) -> str:
    return dt.strftime("%Y%m%d%H%M")


def date_chunks(start: datetime, end: datetime):
    cur = start
    while cur < end:
        nxt = min(cur + timedelta(days=CHUNK_DAYS), end)
        yield cur, nxt
        cur = nxt


def keyword_pattern(kw: str) -> re.Pattern:
    """영문 약어(AI, RPA 등)는 앞뒤가 영문자가 아닐 때만 매칭 → 'MAIL', 'TRAIN' 오탐 방지"""
    if re.fullmatch(r"[A-Za-z0-9]+", kw):
        return re.compile(rf"(?<![A-Za-z]){re.escape(kw)}(?![A-Za-z])", re.IGNORECASE)
    return re.compile(re.escape(kw))


# ───────────────────────── API 호출 ─────────────────────────
class ApiError(Exception):
    pass


def call_api(url: str, params: dict, retries: int = 3) -> tuple[list[dict], int]:
    """한 페이지 호출 → (items, totalCount)"""
    last_err = None
    for attempt in range(retries):
        try:
            r = requests.get(url, params=params, timeout=30)
            text = r.text.strip()
            if text.startswith("<"):
                # 인증키 오류 등은 XML로 옴
                m = re.search(r"<returnAuthMsg>(.*?)</returnAuthMsg>|<resultMsg>(.*?)</resultMsg>", text)
                msg = next((g for g in (m.groups() if m else []) if g), text[:200])
                raise ApiError(f"API 오류 응답: {msg}")
            data = r.json()
            # 인증키 미등록 등 게이트웨이 오류는 JSON(OpenAPI_ServiceResponse)으로도 옴
            gw = data.get("OpenAPI_ServiceResponse")
            if gw:
                h = gw.get("cmmMsgHeader", {}) or {}
                raise ApiError(f"API 오류 응답: {h.get('returnAuthMsg') or h.get('errMsg') or gw}")
            if "response" not in data:
                raise ApiError(f"알 수 없는 응답(HTTP {r.status_code}): {text[:200]}")
            resp = data["response"]
            header = resp.get("header", {})
            code = str(header.get("resultCode", "00"))
            if code not in ("00", "0"):
                raise ApiError(f"[{code}] {header.get('resultMsg')}")
            body = resp.get("body", {}) or {}
            items = body.get("items") or []
            if isinstance(items, dict):
                items = items.get("item") or []
            if isinstance(items, dict):
                items = [items]
            return items, int(body.get("totalCount") or 0)
        except ApiError:
            raise
        except Exception as e:  # 네트워크·JSON 오류는 재시도
            last_err = e
            time.sleep(1.5 * (attempt + 1))
    raise ApiError(f"요청 실패: {last_err}")


def fetch_all(url: str, base_params: dict) -> list[dict]:
    out = []
    for page in range(1, MAX_PAGES + 1):
        params = {**base_params, "pageNo": page, "numOfRows": NUM_ROWS, "type": "json"}
        items, total = call_api(url, params)
        out.extend(items)
        time.sleep(REQUEST_GAP)
        if page * NUM_ROWS >= total or not items:
            break
    return out


# ───────────────────────── 정규화 ─────────────────────────
def norm_bid(raw: dict, bid_type: str, today: str) -> dict:
    no = raw.get("bidNtceNo", "")
    ord_ = raw.get("bidNtceOrd", "") or "000"
    org = raw.get("ntceInsttNm", "")
    demand = raw.get("dminsttNm", "") or org
    return {
        "id": f"{no}-{ord_}",
        "no": no,
        "ord": ord_,
        "type": bid_type,
        "title": (raw.get("bidNtceNm") or "").strip(),
        "org": org,
        "demand": demand,
        "region": OTHER_REGION,   # 분류는 postprocess()에서 config 규칙으로 계산
        "noticeDt": raw.get("bidNtceDt", ""),
        "closeDt": raw.get("bidClseDt", ""),
        "budget": to_int(raw.get("asignBdgtAmt")),
        "estPrice": to_int(raw.get("presmptPrce")),
        "method": raw.get("cntrctCnclsMthdNm", ""),
        "url": raw.get("bidNtceDtlUrl") or raw.get("bidNtceUrl") or "",
        "firstSeen": today,
        "keywords": [],
        "award": None,
    }


def norm_award(raw: dict) -> dict:
    return {
        "no": raw.get("bidNtceNo", ""),
        "winner": raw.get("bidwinnrNm", ""),
        "bizno": raw.get("bidwinnrBizno", ""),
        "amount": to_int(raw.get("sucsfbidAmt")),
        "rate": to_float(raw.get("sucsfbidRate")),
        "openDt": raw.get("rlOpengDt") or raw.get("fnlSucsfDate") or "",
    }


# ───────────────────────── 메인 로직 ─────────────────────────
def collect(cfg: dict, key: str, days: int) -> tuple[list[dict], list[dict], list[str]]:
    now = datetime.now(KST)
    start = now - timedelta(days=days)
    today = now.strftime("%Y-%m-%d")
    base = cfg["api_base"].rstrip("/")
    kw_pats = {kw: keyword_pattern(kw) for kw in cfg["keywords"]}
    rules = Rules(cfg)
    errors: list[str] = []

    bids: dict[str, dict] = {}
    awards: list[dict] = []

    for bid_type in cfg["bid_types"]:
        bid_url = f"{base}/{cfg['endpoints']['bid'][bid_type]}"
        award_url = f"{base}/{cfg['endpoints']['award'][bid_type]}"
        for kw in cfg["keywords"]:
            for s, e in date_chunks(start, now):
                rng = {"inqryBgnDt": fmt_dt(s), "inqryEndDt": fmt_dt(e), "bidNtceNm": kw, "ServiceKey": key}
                # 입찰공고
                try:
                    for raw in fetch_all(bid_url, {**rng, "inqryDiv": "1"}):
                        b = norm_bid(raw, bid_type, today)
                        if not b["no"]:
                            continue
                        if rules.excluded(b["title"]):   # 취소공고, 조류인플루엔자(AI) 등
                            continue
                        matched = [k for k, p in kw_pats.items() if p.search(b["title"])]
                        if not matched:      # 서버는 부분일치라 오탐이 섞임 → 한 번 더 거름
                            continue
                        if b["id"] in bids:
                            bids[b["id"]]["keywords"] = sorted(set(bids[b["id"]]["keywords"]) | set(matched))
                        else:
                            b["keywords"] = matched
                            bids[b["id"]] = b
                except ApiError as ex:
                    errors.append(f"공고/{bid_type}/{kw}/{fmt_dt(s)[:8]}: {ex}")
                # 낙찰정보
                try:
                    for raw in fetch_all(award_url, {**rng, "inqryDiv": cfg.get("award_inqry_div", "2")}):
                        a = norm_award(raw)
                        if a["no"]:
                            awards.append(a)
                except ApiError as ex:
                    errors.append(f"낙찰/{bid_type}/{kw}/{fmt_dt(s)[:8]}: {ex}")
            print(f"  · {bid_type} '{kw}' 완료 (누적 공고 {len(bids)}건, 낙찰 {len(awards)}건)")
    return list(bids.values()), awards, errors


def merge(existing: list[dict], new_bids: list[dict], awards: list[dict]) -> tuple[list[dict], int, int]:
    by_id = {b["id"]: b for b in existing}
    # 중복 정리로 대표 공고에 묶여 들어간 공고 id 도 '이미 아는 공고'로 본다
    known = {i for b in existing for i in (b.get("dupIds") or [b["id"]])}
    added = 0
    for b in new_bids:
        old = by_id.get(b["id"])
        if old:
            b["firstSeen"] = old.get("firstSeen", b["firstSeen"])
            b["award"] = old.get("award")
            b["keywords"] = sorted(set(old.get("keywords", [])) | set(b["keywords"]))
            if old.get("dupIds"):
                b["dupIds"] = old["dupIds"]
        elif b["id"] not in known:
            added += 1
        by_id[b["id"]] = b

    # 낙찰은 공고번호 기준(차수 무관)으로 연결. 묶인 공고의 공고번호로 들어온 낙찰도 대표 공고에 붙인다
    award_by_no = {a["no"]: a for a in awards}
    linked = 0
    for b in by_id.values():
        nos = [b["no"]] + sorted({i.rsplit("-", 1)[0] for i in (b.get("dupIds") or [])} - {b["no"]})
        a = next((award_by_no[n] for n in nos if n in award_by_no), None)
        if a:
            a = {k: v for k, v in a.items() if k != "no"}
            if b.get("award") != a:
                b["award"] = a
                linked += 1
    items = sorted(by_id.values(), key=lambda x: x.get("noticeDt", ""), reverse=True)
    return items, added, linked


def save(cfg: dict, items: list[dict], stats: dict, errors: list[str], dry_run: bool):
    out = {
        "updated": datetime.now(KST).isoformat(timespec="minutes"),
        "keywords": cfg["keywords"],
        "rules": rules_meta(cfg),
        "stats": {"total": len(items), **stats},
        "errors": errors[:20],
        "items": items,
    }
    if dry_run:
        print(json.dumps(items[:3], ensure_ascii=False, indent=2))
        return
    DATA_PATH.parent.mkdir(parents=True, exist_ok=True)
    DATA_PATH.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")


def print_summary(before: list[dict], after: list[dict], pstats: dict, rules: Rules):
    from collections import Counter
    print(f"  · 제외 패턴으로 뺀 공고 {pstats['excluded']}건")
    print(f"  · 중복 정리 {pstats['beforeDedup']}건 → {pstats['afterDedup']}건 "
          f"({pstats['beforeDedup'] - pstats['afterDedup']}건 묶음)")
    kept = [b for b in before if not rules.excluded(b.get("title", ""))]
    c0 = Counter(b.get("region") for b in kept)
    c1 = Counter(rules.region(b.get("demand", ""), b.get("org", "")) for b in kept)
    c2 = Counter(b.get("region") for b in after)
    print("  · 지역별 건수   기존규칙 → 새규칙(중복정리 전) → 최종(중복정리 후)")
    order = rules.region_order
    for r in order + sorted((set(c0) | set(c1) | set(c2)) - set(order)):
        if c0[r] or c1[r] or c2[r]:
            print(f"      {r:<6} {c0[r]:>6} → {c1[r]:>6} ({c1[r] - c0[r]:+4d}) → {c2[r]:>6}")


def main():
    ap = argparse.ArgumentParser(description="나라장터 공고 모니터링 수집기")
    ap.add_argument("--backfill", type=int, help="최근 N일을 몰아서 수집")
    ap.add_argument("--dry-run", action="store_true", help="파일 저장 없이 결과만 출력")
    ap.add_argument("--reprocess", action="store_true",
                    help="API 호출 없이 기존 data/bids.json 에 제외·분류·중복 규칙만 다시 적용")
    args = ap.parse_args()

    cfg = load_config()
    rules = Rules(cfg)

    existing, prev = [], {}
    if DATA_PATH.exists():
        prev = json.loads(DATA_PATH.read_text(encoding="utf-8"))
        existing = prev.get("items", [])

    if args.reprocess:
        if not existing:
            sys.exit("❌ 다시 처리할 data/bids.json 이 없습니다.")
        print(f"▶ 기존 데이터 {len(existing)}건에 규칙 다시 적용 (API 호출 없음)")
        before = [dict(b) for b in existing]
        items, pstats = postprocess(existing, rules)
        print_summary(before, items, pstats, rules)
        stats = {**(prev.get("stats") or {}), **pstats}
        stats.pop("total", None)
        out_updated = prev.get("updated")
        save(cfg, items, stats, prev.get("errors") or [], args.dry_run)
        if not args.dry_run and out_updated:   # 수집 시각은 그대로 둔다
            j = json.loads(DATA_PATH.read_text(encoding="utf-8"))
            j["updated"] = out_updated
            DATA_PATH.write_text(json.dumps(j, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"✔ 전체 {len(items)}건")
        return

    key = get_service_key()
    days = args.backfill or (cfg["backfill_days"] if not existing else cfg["daily_lookback_days"])
    print(f"▶ 최근 {days}일 수집 시작 — 키워드 {len(cfg['keywords'])}개, 유형 {cfg['bid_types']}")

    new_bids, awards, errors = collect(cfg, key, days)
    merged, added, linked = merge(existing, new_bids, awards)

    total_calls = len(cfg["keywords"]) * len(cfg["bid_types"]) * 2 * max(1, -(-days // CHUNK_DAYS))
    if errors and len(errors) >= total_calls:
        # 전부 실패 = 키/승인 문제 → 기존 데이터를 덮어쓰지 않고 실패 처리
        print("\n".join(errors[:5]))
        sys.exit("❌ 모든 호출이 실패했습니다. 인증키 또는 API 활용신청 승인 상태를 확인하세요.")

    items, pstats = postprocess(merged, rules)
    print(f"✔ 신규 {added}건, 낙찰 연결 {linked}건, 전체 {len(items)}건 (오류 {len(errors)}건)")
    print(f"  · 제외 {pstats['excluded']}건, 중복 정리 {pstats['beforeDedup']}→{pstats['afterDedup']}건")
    for e in errors[:5]:
        print("  ⚠", e)
    save(cfg, items, {"addedThisRun": added, "awardsLinked": linked, **pstats}, errors, args.dry_run)


if __name__ == "__main__":
    main()
