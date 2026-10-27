# -*- coding: utf-8 -*-
"""관리 대상 = 사용자의 최신 업무배정표(바탕화면 '업무배정표' 폴더, 매월 1~2회 저장)에 있는 SPC·펀드·리츠만.

업무배정표(엑셀 한 시트) 열은 머리글로 찾는다: 회사명 / 업무(현재: A·T·AT·C) / 종류 / 사업자번호(값 꼴) / 결산월(값 꼴 '12', '3/6/9/12').
- 수집기: 대상이 아닌 SPC 자료는 저장하지 않는다(_미분류는 그대로 - 대상일 수도 있으므로).
- 자동결산: 대상 중 감사·검토 업무(A·AT)만 DSD까지 만든다. 세무만(T)은 건너뜀.
- 결산월: 배정표의 결산월을 그대로 쓴다(추정보다 우선).

이름 맞추기: 공백·(유)·(주)·유동화전문·괄호를 빼고, '제일차'↔'제1차'를 같게 본 뒤, 같거나 한쪽이 다른 쪽을 품으면(4자 이상) 같은 회사.
"""
import os, re, glob, datetime as dt

STRIP = re.compile(r"유동화\s*전문|유한\s*회사|주식\s*회사|\(\s*(주|유)\s*\)|㈜|\s|[()\[\]]")
KN = {k: i + 1 for i, k in enumerate("일이삼사오육칠팔구")}
P_BRN = re.compile(r"^\d{3}-\d{2}-\d{5}$")
P_MONTHS = re.compile(r"^\s*\d{1,2}(\s*/\s*\d{1,2})*\s*$")


def _kor_num(s):
    if s == "십":
        return 10
    if "십" in s:
        a, _, b = s.partition("십")
        return (KN.get(a, 1) if a else 1) * 10 + (KN.get(b, 0) if b else 0)
    return KN.get(s)


def norm(name):
    s = STRIP.sub("", str(name or ""))
    s = re.sub(r"제([일이삼사오육칠팔구십]{1,3})차", lambda m: f"제{_kor_num(m.group(1)) or m.group(1)}차", s)
    return s.lower()


def latest_file(folder):
    fs = [f for f in glob.glob(os.path.join(glob.escape(folder), "**", "*.xls*"), recursive=True)
          if not os.path.basename(f).startswith("~$")]
    return max(fs, key=os.path.getmtime) if fs else None


GENERIC = re.compile(r"(기업구조조정|위탁관리|자기관리)?부동산투자회사|투자회사|전문투자형|일반|사모|투자신탁|증권|혼합자산|리츠")
KNUM = re.compile(r"[일이삼사오육칠팔구십백천]+(?=호|차)")


def _kor_big(s):
    total, cur = 0, 0
    for ch in s:
        if ch in KN:
            cur = KN[ch]
        elif ch in "십백천":
            total += (cur or 1) * {"십": 10, "백": 100, "천": 1000}[ch]
            cur = 0
    return total + cur


def loose(name):
    """느슨한 키(10-24): 펀드·리츠 이름의 일반 단어(일반·전문투자형·사모·투자신탁·위탁관리부동산투자회사 등) 빼고, '일백이십오호'→125호, 제·호·차 빼기"""
    s = norm(name)
    s = KNUM.sub(lambda m: str(_kor_big(m.group(0))), s)
    s = GENERIC.sub("", s)
    return re.sub(r"[제호차]", "", s)


def _digits(s):
    return re.findall(r"\d+", s)


class Scope:
    def __init__(self, entries, path):
        self.entries, self.path = entries, path
        self._cache = {}

    def match(self, name):
        """SPC·폴더 이름 → 배정표 항목(dict) 또는 None"""
        if not name:
            return None
        if name in self._cache:
            return self._cache[name]
        n = norm(name)
        hit = None
        for e in self.entries:
            k = e["키"]
            if n == k or (len(k) >= 4 and k in n) or (len(n) >= 4 and n in k):
                hit = e
                break
        if hit is None:                                  # 펀드·리츠: 일반 단어·운용사 접두어·한글 숫자·오타 차이(10-24 전수 시험)
            import difflib
            ln = loose(name)
            for e in self.entries:
                lk = e.setdefault("느슨", loose(e["이름"]))
                if len(ln) < 4 or len(lk) < 4 or _digits(ln) != _digits(lk):
                    continue
                short, long_ = sorted((ln, lk), key=len)
                if ln == lk or (len(short) >= 6 and long_.endswith(short)) or \
                        (min(len(ln), len(lk)) >= 8 and difflib.SequenceMatcher(None, ln, lk).ratio() >= 0.92):
                    hit = e
                    break
        self._cache[name] = hit
        return hit

    def __len__(self):
        return len(self.entries)


def load(cfg):
    """설정 assignment_dir(기본 바탕화면\\업무배정표)의 최신 배정표 → Scope 또는 None(없거나 scope: off)"""
    if str(cfg.get("scope", "assignment")).lower() in ("off", "none", "false", "all"):
        return None
    folder = cfg.get("assignment_dir") or os.path.join(os.path.dirname(os.path.normpath(cfg.get("spc_root") or "")), "업무배정표")
    p = latest_file(folder) if os.path.isdir(folder) else None
    if not p:
        return None
    import openpyxl
    ws = openpyxl.load_workbook(p, read_only=True, data_only=True).worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    if not rows:
        return None
    head = [str(h or "").replace(" ", "") for h in rows[0]]
    col = lambda *names: next((i for i, h in enumerate(head) if any(n in h for n in names)), None)
    c_name, c_work, c_kind = col("회사명"), col("업무(현재)", "업무"), col("종류")
    body = [r for r in rows[1:] if c_name is not None and len(r) > c_name and r[c_name]]
    def by_values(rx):
        best, n = None, 0
        for i in range(len(head)):
            k = sum(1 for r in body if i < len(r) and r[i] is not None and rx.match(str(r[i])))
            if k > n:
                best, n = i, k
        return best if n >= max(3, len(body) // 3) else None
    c_brn, c_mon = by_values(P_BRN), by_values(P_MONTHS)
    entries = []
    for r in body:
        g = lambda c: (r[c] if c is not None and c < len(r) else None)
        months = [int(x) for x in re.findall(r"\d{1,2}", str(g(c_mon) or "")) if 1 <= int(x) <= 12]
        if "매월" in str(g(c_mon) or ""):
            months = list(range(1, 13))
        entries.append({"이름": str(g(c_name)).strip(), "키": norm(g(c_name)), "사업자번호": g(c_brn),
                        "업무": str(g(c_work) or "").strip().upper(), "종류": g(c_kind), "결산월": sorted(set(months))})
    return Scope(entries, p)


def needs_report(entry):
    """DSD(감사·검토보고서)까지 만들 대상인가 - 업무에 A(감사·검토)가 있으면"""
    return bool(entry) and "A" in (entry.get("업무") or "")
