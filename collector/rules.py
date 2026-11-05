"""
규칙 기반 분류·정리 (외부 LLM 없음)

config.json 의 규칙으로
  - 제외 패턴 적용 (AI 무관 공고: 조류인플루엔자 등)
  - 지역(시도) 추정
  - 기관구분(kind) 추정
  - 분야(field) 분류
  - 재공고·정정공고 중복 정리(dups)
를 계산한다. 수집기와 화면이 같은 기준을 쓰도록 결과를 bids.json 에 저장한다.
"""
from __future__ import annotations

import re

OTHER_REGION = "중앙·기타"


class Rules:
    def __init__(self, cfg: dict):
        self.exclude = [re.compile(p, re.IGNORECASE) for p in cfg.get("exclude_patterns", [])]
        self.exclude_kw = [e.lower() for e in cfg.get("exclude_keywords", [])]

        rr = cfg.get("region_rules", {})
        self.region_order = rr.get("order") or []
        self.region_overrides = [(re.compile(p), r) for p, r in (rr.get("overrides") or {}).items()]
        self.ignore_org = tuple(rr.get("ignore_org_prefixes") or ())
        self.strip_prefixes = sorted(rr.get("strip_prefixes") or [], key=len, reverse=True)
        # 정식명칭은 긴 것부터 비교 (예: '전남광주통합특별시'가 '전남'보다 먼저)
        self.region_official = sorted(
            ((name, region) for region, names in (rr.get("official") or {}).items() for name in names),
            key=lambda x: len(x[0]), reverse=True)
        self.region_patterns = [(re.compile(p), region)
                                for region, pats in (rr.get("patterns") or {}).items() for p in pats]

        self.field_rules = [(label, re.compile(p, re.IGNORECASE)) for label, p in cfg.get("field_rules", [])]
        fd = cfg.get("field_default") or {}
        self.field_default = fd
        self.kind_rules = [(label, re.compile(p)) for label, p in cfg.get("kind_rules", [])]
        self.kind_default = cfg.get("kind_default", "기타")
        self.dedup_strip = re.compile(cfg.get("dedup_strip_pattern") or r"(?!x)x")

    # ───────── 제외 ─────────
    def excluded(self, title: str) -> bool:
        t = title or ""
        if any(x in t.lower() for x in self.exclude_kw):
            return True
        return any(p.search(t) for p in self.exclude)

    # ───────── 지역 ─────────
    def _strip(self, name: str) -> str:
        n = name.strip()
        changed = True
        while changed:
            changed = False
            for p in self.strip_prefixes:
                if n.startswith(p):
                    n = n[len(p):].lstrip()
                    changed = True
        return n

    def _region_of(self, name: str) -> str | None:
        if not name:
            return None
        for pat, region in self.region_overrides:
            if pat.search(name):
                return region
        n = self._strip(name)
        for official, region in self.region_official:
            if n.startswith(official):
                return region
        # 기관명에서 가장 앞에 나오는 지명 (같은 위치면 규칙 순서)
        best = None
        for i, (pat, region) in enumerate(self.region_patterns):
            m = pat.search(n)
            if m and (best is None or (m.start(), i) < best[0]):
                best = ((m.start(), i), region)
        return best[1] if best else None

    def region(self, demand: str, org: str) -> str:
        r = self._region_of(demand)
        if r:
            return r
        if org and org != demand and not org.startswith(self.ignore_org):
            r = self._region_of(org)
        return r or OTHER_REGION

    # ───────── 기관구분 ─────────
    def kind(self, demand: str) -> str:
        n = (demand or "").strip()
        for label, pat in self.kind_rules:
            if pat.search(n):
                return label
        return self.kind_default

    # ───────── 분야 ─────────
    def field(self, title: str, bid_type: str) -> str:
        for label, pat in self.field_rules:
            if pat.search(title or ""):
                return label
        return self.field_default.get(bid_type) or self.field_default.get("*", "기타")

    def classify(self, b: dict) -> dict:
        b["region"] = self.region(b.get("demand", ""), b.get("org", ""))
        b["kind"] = self.kind(b.get("demand") or b.get("org") or "")
        b["field"] = self.field(b.get("title", ""), b.get("type", ""))
        return b

    # ───────── 중복 정리 ─────────
    def dedup_key(self, b: dict) -> tuple[str, str]:
        t = self.dedup_strip.sub(" ", b.get("title") or "")
        t = re.sub(r"\s+", " ", t).strip(" -_·")
        return t, (b.get("demand") or "").strip()

    def dedup(self, items: list[dict]) -> list[dict]:
        """공고명+수요기관 또는 공고번호가 같은 공고를 1건으로 묶는다.
        대표: 낙찰 결과 있는 공고 → 없으면 최신 공고일. dupIds 에 묶인 공고 id 를 모두 남긴다."""
        parent = list(range(len(items)))

        def find(i):
            while parent[i] != i:
                parent[i] = parent[parent[i]]
                i = parent[i]
            return i

        def union(a, b):
            ra, rb = find(a), find(b)
            if ra != rb:
                parent[rb] = ra

        seen: dict = {}
        for i, b in enumerate(items):
            for key in (("k",) + self.dedup_key(b), ("no", b.get("no"))):
                if key in seen:
                    union(seen[key], i)
                else:
                    seen[key] = i

        groups: dict[int, list[dict]] = {}
        for i, b in enumerate(items):
            groups.setdefault(find(i), []).append(b)

        out = []
        for members in groups.values():
            rep = max(members, key=lambda x: (x.get("award") is not None, x.get("noticeDt") or "", x.get("id") or ""))
            rep = dict(rep)
            ids = set()
            for m in members:
                ids.update(m.get("dupIds") or [m["id"]])
            rep["dupIds"] = sorted(ids)
            rep["dups"] = len(ids)
            rep["keywords"] = sorted({k for m in members for k in (m.get("keywords") or [])})
            seen_dates = [m.get("firstSeen") for m in members if m.get("firstSeen")]
            if seen_dates:
                rep["firstSeen"] = min(seen_dates)
            out.append(rep)
        out.sort(key=lambda x: x.get("noticeDt", ""), reverse=True)
        return out


def postprocess(items: list[dict], rules: Rules) -> tuple[list[dict], dict]:
    """제외 → 분류 → 중복 정리. (결과, 통계) 반환"""
    kept = [b for b in items if not rules.excluded(b.get("title", ""))]
    stats = {"excluded": len(items) - len(kept)}
    for b in kept:
        rules.classify(b)
    before = len({i for b in kept for i in (b.get("dupIds") or [b["id"]])})
    out = rules.dedup(kept)
    stats.update({"beforeDedup": before, "afterDedup": len(out)})
    return out, stats


def rules_meta(cfg: dict) -> dict:
    """화면·엑셀·보고서가 같은 기준을 쓰도록 bids.json 에 함께 저장하는 규칙 요약"""
    rr = cfg.get("region_rules", {})
    fields = []
    for label, _ in cfg.get("field_rules", []):
        if label not in fields:
            fields.append(label)
    for v in (cfg.get("field_default") or {}).values():
        if v not in fields:
            fields.append(v)
    kinds = []
    for label, _ in cfg.get("kind_rules", []):
        if label not in kinds:
            kinds.append(label)
    if cfg.get("kind_default", "기타") not in kinds:
        kinds.append(cfg.get("kind_default", "기타"))
    return {
        "homeRegion": cfg.get("home_region", "제주"),
        "regions": rr.get("order") or [],
        "fields": fields,
        "fieldRules": cfg.get("field_rules", []),
        "kinds": kinds,
        "excludePatterns": cfg.get("exclude_patterns", []),
        "statsExcludeFields": cfg.get("stats_exclude_fields", []),
        "similarExcludeFields": cfg.get("similar_check_exclude_fields", []),
        "dedup": "공고명(공백 정리·재공고 표시 제거)+수요기관이 같거나 공고번호가 같으면 1건. 낙찰 결과 있는 공고 우선, 없으면 최신 공고일",
    }
