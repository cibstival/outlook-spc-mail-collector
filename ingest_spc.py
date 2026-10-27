# -*- coding: utf-8 -*-
"""SPC 자료 메일 자동 수집·분류 - 자금판·거래내역이 메일로 오면 SPC별·결산월별 폴더에 넣는다.

자동화 목표(가이드1)의 마지막 고리: "셋팅 후에는 자금판·거래내역만 추가하면 A·B·D·DSD가 나온다"에서
'추가'를 사람 대신 한다. 원형은 이전 감사조서 수집 스크립트(편지함 기준·EntryID 중복 방지·
zip 풀기·대용량 링크 표시)이고, 분류 기준을 SPC용으로 바꿨다.

흐름
  메일(편지함 또는 .eml/.msg 폴더) → 첨부 저장 → zip 풀기 → 파일마다
     ① 수탁사: 보낸 사람 도메인(설정 domains) → 없으면 SPC로 역추적
     ② SPC  : 제목·파일명에 든 SPC 이름(설정 aliases + SPC 루트 폴더 이름에서 자동 생성한 별칭)
     ③ 결산월: 파일명·제목의 날짜 중 가장 늦은 것(기간 표기 '~'면 끝 날짜) → YY.MM
     ④ 종류  : 자금판 / 거래내역 / 잔액증명 / 발행내역 / 운용지시서 / 세금계산서 / 기타
  → <dest_root>/<수탁사>/<SPC>/<종류 폴더>/<YY.MM>/파일   (못 정하면 <dest_root>/_미분류/)
  → _수신로그.md, _새자료.json(생성 파이프라인이 읽는 대기열)

지키는 것
  - 메일 본문은 저장하지 않는다(메일 분석 키트 원칙). 제목·보낸이·첨부 이름만 로그에.
  - 기존 파일은 덮어쓰지 않는다(같은 이름이면 (1),(2)…). 원본 SPC 폴더를 직접 대상으로 쓰지 않는 것이 기본값.
  - 같은 메일은 두 번 처리하지 않는다(Outlook EntryID 또는 Message-ID).

사용
  python ingest_spc.py --config spc_mail_config.json            Outlook(Windows 클래식) 메일 폴더 전부에서 SPC 자료 메일만
  python ingest_spc.py --config ... --survey [--days 180]       편지함 조사(저장 없음): 폴더별 SPC 자료 메일 수·모르는 도메인
  python ingest_spc.py --config ... --eml <폴더>                .eml 파일 폴더(시험·다른 메일 프로그램)
  python ingest_spc.py --config ... --dry-run                   저장하지 않고 분류만 출력
  python ingest_spc.py --config ... --baseline                  지금 있는 메일을 '처리완료'로만 표시
  python ingest_spc.py --config ... --dry-run --days 60         최근 60일 메일만 (시험용)
  python ingest_spc.py --config ... --days 400 --no-queue        과거 메일 한 번 가져오기(자동결산 출발점용, 대기열에는 안 올림)
  python ingest_spc.py --config ... --reroute                   _미분류 파일을 지금 규칙(파일 내용 포함)으로 다시 분류
  설정 exclude_folders: SPC 루트에서 수탁사가 아닌 작업 폴더 이름 목록(그 아래는 SPC로 보지 않음)
"""
import filecmp
import os, re, sys, json, shutil, zipfile, datetime as dt, traceback, email, email.policy, hashlib
from email.utils import parseaddr, parsedate_to_datetime

HERE = os.path.dirname(os.path.abspath(__file__))

KINDS = [   # (종류, 넣을 하위 폴더, 파일명·제목 규칙) - 위에서부터 먼저 걸리는 것
    ("운용지시서", "운용지시서", r"운용\s*지시|자금\s*운용|지시서"),
    ("세금계산서", "세금계산서", r"세금\s*계산서|전자\s*세금|계산서"),
    ("잔액증명", "자금판 및 거래내역", r"잔액\s*증명|잔고\s*증명|금융\s*거래\s*조회|조회서"),
    ("발행내역", "자금판 및 거래내역", r"발행\s*내역|어음\s*발행|발행\s*현황|사채\s*원리금|원리금\s*상환"),
    ("거래내역", "자금판 및 거래내역", r"거래\s*내역|입출금|통장|계좌|거래명세|일일\s*거래|사업별\s*입출금"),
    ("자금판", "자금판 및 거래내역", r"자금\s*판|자금\s*관리|자금\s*처리|자금\s*수지|비용표|(?<![A-Za-z])CF(?![A-Za-z])|cash\s*flow|"
                                 r"사후\s*관리|자산\s*관리\s*월보|월보|현금\s*흐름"),
]
# 기장 대행(1차 결산) 쪽이 보내는 작성자 조서·보고서 - 파일 이름만 본다(엑셀 A·B·D, DSD)
AUTHOR_KINDS = [
    ("작성자A", "작성자 A·B", r"^A[\s.\-_]|정산표", (".xls", ".xlsx", ".xlsm")),
    ("작성자B", "작성자 A·B", r"^B[\s.\-_]|결산보정|기말조정", (".xls", ".xlsx", ".xlsm")),
    ("작성자D", "작성자 A·B", r"^D[\s.\-_]|리드시트", (".xls", ".xlsx", ".xlsm")),
    ("DSD", "보고서", r".", (".dsd",)),
]
QUEUE_KINDS = ("자금판", "거래내역", "발행내역", "잔액증명", "작성자A", "작성자B", "DSD")   # 자동결산 대기열에 올릴 종류
PREPARER_DOMAINS = {"preparer.example.kr": "기장대행사(예시)"}   # 1차 기장·A·B 작성(설정 preparer_domains로 추가)
INTERNAL_DOMAINS = {"ourfirm.example.kr"}                  # 우리 법인(우리 법인) - 동료가 전달한 메일은 원래 보낸 곳으로 판단(설정 internal_domains)
FWD_FROM = re.compile(r"(?:^|\n)\s*(?:From|보낸\s*사람|보낸사람|발신자)\s*[:：]\s*[^\n]*?([\w.+\-]+@[\w\-]+(?:\.[\w\-]+)+)", re.I)


def original_sender(sender, body, cfg):
    """우리 법인 사람이 보낸(전달·회신) 메일이면 본문의 전달 머리글에서 처음 나오는 바깥 보낸이. 본문은 저장하지 않는다."""
    internal = INTERNAL_DOMAINS | {d.lower() for d in (cfg.get("internal_domains") or [])}
    dom = sender.split("@")[-1].lower() if "@" in sender else ""
    if dom not in internal:
        return sender
    for m in FWD_FROM.finditer(body or ""):
        addr = m.group(1)
        if addr.split("@")[-1].lower() not in internal:
            return addr
    return sender
SKIP_EXT = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".emz", ".ico"}
SKIP_NAME = re.compile(r"^(~WRD\d+|image\d+|Outlook-|oledata)", re.I)
PERIOD_DIR = re.compile(r"^\s*(19|20)\d{2}([\s._\-년]|$)|^\s*\d{2}[.\-]\d{2}\s*$")   # '2025.12', '2025 1기 확정' 같은 기간 폴더는 SPC가 아님
NEW_SPC = re.compile(r"신보\s*\d{4}\s*-?\s*제?\s*\d+\s*차|(?<![가-힣A-Za-z])[가-힣A-Za-z][가-힣A-Za-z0-9]{1,13}?제\s*[일이삼사오육칠팔구십\d]{1,3}\s*차"
                     r"|(?<![가-힣A-Za-z])[가-힣A-Za-z][가-힣A-Za-z0-9]{1,13}?(?=\s*유동화\s*전문)")
BANK_DOMAINS = {   # 공개된 은행 메일 도메인 → 은행 이름(설정 domains가 우선)
    "ibk.co.kr": "기업은행", "kbstar.com": "국민은행", "kbfg.com": "국민은행", "shinhan.com": "신한은행",
    "wooribank.com": "우리은행", "woorifg.com": "우리은행", "hanafn.com": "하나은행", "hanabank.com": "하나은행",
    "nonghyup.com": "농협은행", "nhbank.com": "농협은행", "kdb.co.kr": "산업은행", "suhyup-bank.com": "수협은행",
}
SINBO_NAME = re.compile(r"신보\s*\d{4}\s*제\s*\d+\s*차")
SINBO_ID = re.compile(r"(?<![\d.])(20[12]\d)\s*-\s*(\d{1,2})(?![\d\-])(?!\.\d)")   # 신보 SPC 약칭 '2099-21' = 신보2099제21차


def sinbo_id(text, context):
    """메일·파일에 '신보'가 있을 때만 '2099-21 거래내역' 꼴의 앞부분을 신보 SPC 이름으로 본다(날짜와 구분)."""
    if "신보" not in (context or "") and "SBAID" not in (context or "").upper():
        return None
    m = SINBO_ID.search(text or "")
    return f"신보{m.group(1)}제{int(m.group(2))}차" if m else None


STRIP = re.compile(r"유동화\s*전문|유한\s*회사|주식\s*회사|\(\s*(주|유)\s*\)|㈜|\s")


# ---------------------------------------------------------------- 공통
def safe(s, n=60):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]', " ", s or "")
    return re.sub(r"\s+", " ", s).strip()[:n].rstrip(" .")


def uniq(path):
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(path)
    k = 1
    while os.path.exists(f"{stem}({k}){ext}"):
        k += 1
    return f"{stem}({k}){ext}"


def is_junk(name, size=None):
    base = os.path.basename(name)
    if SKIP_NAME.match(base):
        return True
    return os.path.splitext(base)[1].lower() in SKIP_EXT and (size is None or size < 300_000)


def fix_korean(name):
    try:
        return name.encode("cp437").decode("cp949")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return name


SHEET_EXT = (".xls", ".xlsx", ".xlsm", ".csv", ".txt", ".htm", ".html")


def peek_text(path, max_rows=40, max_cols=20, limit=6000):
    """엑셀·CSV 앞부분 글자(계좌명·SPC명 찾기용). 읽지 못하면 빈 문자열. 내용은 판별에만 쓰고 저장·출력하지 않는다."""
    ext = os.path.splitext(path)[1].lower()
    out = []
    try:
        if ext in (".xlsx", ".xlsm"):
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            try:
                for ws in wb.worksheets[:3]:
                    out.append(ws.title)
                    for row in ws.iter_rows(max_row=max_rows, max_col=max_cols, values_only=True):
                        out += [str(v) for v in row if v not in (None, "")]
            finally:
                wb.close()                         # 열린 채로 두면 옮길 때 'WinError 32 사용 중'(10-27 재분류 중단)
        elif ext == ".xls":
            try:
                import xlrd
                wb = xlrd.open_workbook(path, on_demand=True)
                try:
                    for i in range(min(3, wb.nsheets)):
                        sh = wb.sheet_by_index(i)
                        out.append(sh.name)
                        for r in range(min(max_rows, sh.nrows)):
                            out += [str(v) for v in sh.row_values(r)[:max_cols] if v not in (None, "")]
                finally:
                    wb.release_resources()
            except Exception:                      # 은행이 주는 .xls 중 실제로는 HTML·텍스트인 것
                raw = open(path, "rb").read(200_000)
                for enc in ("utf-8", "cp949"):
                    try:
                        out.append(re.sub(r"<[^>]+>", " ", raw.decode(enc)))
                        break
                    except UnicodeDecodeError:
                        continue
        elif ext in (".csv", ".txt", ".htm", ".html"):
            raw = open(path, "rb").read(200_000)
            for enc in ("utf-8-sig", "cp949"):
                try:
                    out.append(re.sub(r"<[^>]+>", " ", raw.decode(enc)))
                    break
                except UnicodeDecodeError:
                    continue
    except Exception:
        return ""
    return " ".join(out)[:limit]


ZIP_PASSWORDS = []          # zip_passwords.txt(로컬 전용, 저장소 제외)에서 읽음. 화면·로그에 내지 않는다


def load_zip_passwords(cfg):
    p = cfg.get("zip_password_file") or os.path.join(HERE, "zip_passwords.txt")
    if os.path.exists(p):
        for enc in ("utf-8-sig", "cp949", "utf-16"):
            try:
                ZIP_PASSWORDS[:] = [x.strip() for x in open(p, encoding=enc).read().splitlines() if x.strip()]
                break
            except (UnicodeDecodeError, UnicodeError):
                continue


def _open_zip(zip_path):
    """암호 zip이면 등록된 비밀번호를 차례로 시도. (zip 객체, 비밀번호 bytes|None) 또는 (None, None)."""
    z = zipfile.ZipFile(zip_path)
    enc = [i for i in z.infolist() if i.flag_bits & 0x1 and not i.is_dir()]
    if not enc:
        return z, None
    name = enc[0].filename
    for pw in ZIP_PASSWORDS:
        for b in {pw.encode("utf-8"), pw.encode("cp949", "ignore")}:
            for aes in (False, True):
                try:
                    if aes:                          # AES 암호는 pyzipper가 있어야 풀림
                        import pyzipper
                        zz = pyzipper.AESZipFile(zip_path)
                    else:
                        zz = z
                    with zz.open(name, pwd=b) as f:
                        f.read(16)
                    return zz, b
                except Exception:
                    continue
    z.close()
    return None, None


def extract_zip(zip_path, dest, depth=0):
    out = []
    if depth > 3:
        return out
    try:
        z, pwd = _open_zip(zip_path)
        if z is None:
            print(f"    ! 암호 zip(맞는 비밀번호 없음, zip 그대로 둠): {os.path.basename(zip_path)}")
            return [zip_path]
        with z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                raw = info.filename if info.flag_bits & 0x800 else fix_korean(info.filename)
                parts = [safe(p, 50) for p in raw.replace("\\", "/").split("/") if p not in ("", ".", "..")]
                if not parts:
                    continue
                target = os.path.join(dest, *parts)
                if len(os.path.abspath(target)) > 230:     # Windows 경로 길이 제한(260자, 10-23 WinError 206) - zip 안 폴더를 빼고 이름만
                    target = os.path.join(dest, safe(parts[-1], 60))
                    if len(os.path.abspath(target)) > 230:
                        stem, ext = os.path.splitext(parts[-1])
                        target = os.path.join(dest, safe(stem, 20) + ext[:10])
                target = uniq(target)
                os.makedirs(os.path.dirname(target), exist_ok=True)
                with z.open(info, pwd=pwd) as src, open(target, "wb") as dst:
                    shutil.copyfileobj(src, dst)
                out += extract_zip(target, os.path.splitext(target)[0], depth + 1) if target.lower().endswith(".zip") else [target]
    except (zipfile.BadZipFile, OSError, RuntimeError, NotImplementedError) as e:
        print(f"    ! zip 실패(zip 그대로 둠): {os.path.basename(zip_path)} ({e})")
        return out or [zip_path]
    return out


SHORT_NAME = re.compile(r"^.+?제?[0-9일이삼사오육칠팔구십백]+(호|차)(?![0-9])")

# ---------------------------------------------------------------- 판별
class Router:
    def __init__(self, cfg):
        self.cfg = cfg
        self.domains = {d.lower(): t for d, t in cfg.get("domains", {}).items()}     # 도메인 → 수탁사
        self.spc_home = {}                                                            # SPC → 수탁사
        self.trustees = set()                                                         # SPC 루트의 수탁사 폴더 이름
        self.alias = []                                                               # (별칭, SPC)
        root = cfg.get("spc_root")
        skip = set(cfg.get("exclude_folders", []))                                    # 수탁사가 아닌 작업 폴더
        if root and os.path.isdir(root):                                               # SPC 루트: <수탁사>/<SPC>/
            for t in sorted(os.listdir(root)):
                p = os.path.join(root, t)
                if os.path.isdir(p) and not t.startswith(("_", ".")) and t not in skip:
                    self.trustees.add(t)
                    for s in sorted(os.listdir(p)):
                        if os.path.isdir(os.path.join(p, s)) and not s.startswith(("_", ".")) \
                                and not PERIOD_DIR.match(s):
                            self.spc_home.setdefault(s, t)
        for s, t in cfg.get("spc", {}).items():                                       # 설정이 폴더보다 우선
            self.spc_home[s] = t
        for s in self.spc_home:
            self.alias.append((STRIP.sub("", s), s))
        for a, s in cfg.get("aliases", {}).items():
            if re.fullmatch(r"신보\d{4}", STRIP.sub("", a)):   # '신보2025'처럼 해만 있는 별칭은 그해 신보 SPC 전부에 걸려서 뺌(10-23)
                continue
            self.alias.append((STRIP.sub("", a), s))
        short = {}                                                                    # 짧은 이름(…제5호·…제일차까지) - 메일 제목은 '[대림제5호]'처럼 줄여 씀(10-27: 대상 메일이 _미분류로 빠지던 문제)
        for s in self.spc_home:
            m = SHORT_NAME.match(STRIP.sub("", s))
            if m and m.group(0) != STRIP.sub("", s) and len(m.group(0)) >= 4 and re.search(r"[가-힣]{2}", m.group(0)):
                short.setdefault(m.group(0), set()).add(s)
        full = {a for a, _ in self.alias}
        for a, ss in short.items():
            if len(ss) == 1 and a not in full and not any(a in f and f not in {STRIP.sub("", x) for x in ss} and f.startswith(a) for f in full):
                self.alias.append((a, ss.pop()))                                       # 두 SPC가 같은 짧은 이름이면 안 씀
        self.alias = sorted({x for x in self.alias if len(x[0]) >= 3}, key=lambda x: -len(x[0]))   # 긴 이름 먼저

    def spc(self, *texts):
        t = STRIP.sub("", " ".join(x or "" for x in texts))
        for a, s in self.alias:
            if a in t:
                return s
        return None

    def spc_all(self, text):
        t = STRIP.sub("", text or "")
        found = {s for a, s in self.alias if a in t}
        found |= {STRIP.sub("", m.group(0)) for m in NEW_SPC.finditer(text or "")} - {STRIP.sub("", x) for x in found}
        return found

    @staticmethod
    def new_spc(*texts):
        """SPC 루트에 폴더가 없는 SPC도 이름 꼴로 알아본다(예: 신보2023제23차, ○○제일차, ○○유동화전문).
        수집함에만 새 폴더를 만들고 로그에 '새 SPC'로 표시한다."""
        for t in texts:
            t = t or ""
            m = NEW_SPC.search(t)
            if m:
                return STRIP.sub("", m.group(0))
        return None

    def trustee(self, sender, spc):
        dom = sender.split("@")[-1].lower() if "@" in sender else ""
        for d, t in self.domains.items():
            if dom == d or dom.endswith("." + d):
                return t
        return self.spc_home.get(spc) if spc else None

    def spc_place(self, spc, per, kind):
        """원래 SPC 폴더에도 둘 자리 → 폴더 경로 또는 None.
        사용자가 예전에 메일 첨부를 직접 저장한 위치에서 배운 규칙(spc_rules.json)이 '확실'하고, 그 SPC가 허용 목록
        (설정 place_allow)에 있고, 그 기간 폴더가 이미 있을 때만. 수집함 사본은 그대로 둔다."""
        if not self.cfg.get("place_in_spc", False) or not spc or spc not in self.spc_home or not per:
            return None
        tru = self.spc_home[spc]
        allow = set(self.cfg.get("place_allow") or [])
        if f"{tru}/{spc}" not in allow and spc not in allow:
            return None
        if not hasattr(self, "_rules"):
            rp = self.cfg.get("place_rules") or os.path.join(HERE, "spc_rules.json")
            try:
                self._rules = json.load(open(rp, encoding="utf-8"))["규칙"]
            except Exception:
                self._rules = {}
        r = self._rules.get(f"{tru}/{spc}|{kind}")
        if not r or not r.get("확실"):
            return None
        sc = getattr(self, "scope", None)
        ent = sc.match(spc) if sc else None
        import spc_rules
        if ent and ent.get("결산월") and kind in spc_rules.MONTHLY_KINDS:   # 매달 자료는 결산월 기준 틀
            per = closing_of(per.rstrip("*"), ent["결산월"])
        import spc_rules
        sd = os.path.join(self.cfg["spc_root"], tru, spc)
        return spc_rules.target_dir(sd, r["틀"], per.rstrip("*"))

    def inbox_home(self, spc):
        """수집함에 이미 이 SPC 폴더가 있으면 그 수탁사·은행 폴더(같은 SPC를 한곳에 모음)."""
        dest = self.cfg.get("dest_root")
        if not spc or not dest or not os.path.isdir(dest):
            return None
        key = STRIP.sub("", spc)
        for t in sorted(os.listdir(dest)):
            if t.startswith("_") and t != "_수탁사미상":
                continue
            tp = os.path.join(dest, t)
            if os.path.isdir(tp) and any(STRIP.sub("", s) == key for s in os.listdir(tp)):
                if t != "_수탁사미상":
                    return t
        return None

    def bank_folder(self, sender, spc):
        """SPC 루트에 없는 SPC는 보낸 은행별로 나눈다. 신보 SPC는 'SBAID-<은행>' 폴더가 있으면 그 이름을 쓴다."""
        dom = sender.split("@")[-1].lower() if "@" in sender else ""
        bank = next((b for d, b in BANK_DOMAINS.items() if dom == d or dom.endswith("." + d)), None)
        if not bank:
            return "_수탁사미상"
        short = bank.replace("은행", "")
        if (spc or "").startswith("신보") and f"SBAID-{short}" in self.trustees:
            return f"SBAID-{short}"
        return bank

    def closing(self, spc, per):
        """설정 closing_months(SPC별 결산월 목록, 예 [6, 12])가 있으면 그 달 이상 가장 가까운 결산월로 올린다
        (5월 27일자 자금판 → 6월 결산 자료)."""
        guessed = per.endswith("*")                  # 받은 날로 추정한 결산월 → 가까운 과거 결산월로 내림
        per = per.rstrip("*")
        months = (self.cfg.get("closing_months") or {}).get(spc)
        if not months:
            return per
        y, m = 2000 + int(per[:2]), int(per[3:])
        for k in range(13):
            d = -k if guessed else k
            mm = (m - 1 + d) % 12 + 1
            yy = y + (m - 1 + d) // 12
            if mm in months:
                return f"{yy % 100:02d}.{mm:02d}"
        return per

    @staticmethod
    def kind(*texts):
        base = os.path.basename(texts[0] or "") if texts else ""
        ext = os.path.splitext(base)[1].lower()
        for k, folder, pat, exts in AUTHOR_KINDS:          # 작성자 A·B·D·DSD는 파일 이름만으로
            if ext in exts and re.search(pat, base):
                return k, folder
        t = " ".join(x or "" for x in texts)
        for k, folder, pat in KINDS:
            if re.search(pat, t, re.I):
                return k, folder
        return "기타", "기타"

    @staticmethod
    def period(texts, received):
        """결산 년월(YY.MM). 기간 '~'의 끝 날짜 > 단일 날짜 중 가장 늦은 것. 받은 날 이후 날짜는 무시(만기일 등)."""
        for t in texts:                             # 부가세 'YYYY년 N기 확정/예정' → 과세기간 끝 달
            m = re.search(r"(20\d{2})\s*[년.]?\s*([12])\s*기\s*(확정|예정)?", t or "")
            if m:
                y, k, w = int(m.group(1)), int(m.group(2)), m.group(3)
                mo = {(1, "예정"): 3, (2, "예정"): 9}.get((k, w), 6 if k == 1 else 12)
                return f"{y % 100:02d}.{mo:02d}"
        cands = []
        for t in texts:
            t = t or ""
            for m in re.finditer(r"(20\d{2})[.\-_/년 ]?\s?(\d{1,2})[.\-_/월 ]?\s?(\d{1,2})?일?", t):
                y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3) or 1)
                cands.append((y, mo, d))
            for m in re.finditer(r"(?<!\d)(\d{2})(\d{2})(\d{2})(?!\d)", t):          # 260627 꼴
                y, mo, d = 2000 + int(m.group(1)), int(m.group(2)), int(m.group(3))
                cands.append((y, mo, d))
            for m in re.finditer(r"(?<!\d)(\d{2})\.(\d{2})(?!\d)(?!\.\d)", t):       # 26.06 꼴
                cands.append((2000 + int(m.group(1)), int(m.group(2)), 1))
            for m in re.finditer(r"(?<=[_\s(])(\d{2})(\d{2})(?=[_.)\s]|$)", t):       # _2606. 꼴(YYMM)
                y, mo = int(m.group(1)), int(m.group(2))
                if 15 <= y <= 40 and 1 <= mo <= 12:
                    cands.append((2000 + y, mo, 1))
        ok = []
        for y, mo, d in cands:
            try:
                q = dt.date(y, mo, max(1, min(d, 28)))
            except ValueError:
                continue
            if dt.date(2015, 1, 1) <= q <= received.date() + dt.timedelta(days=31):
                ok.append(q)
        if ok:
            q = max(ok)
            return f"{q.year % 100:02d}.{q.month:02d}"
        q = received.date().replace(day=1) - dt.timedelta(days=1)          # 날짜가 없으면 받은 달의 전달(결산 후 송부)
        return f"{q.year % 100:02d}.{q.month:02d}*"


# ---------------------------------------------------------------- 메일 원천
class Mail:
    def __init__(self, mid, received, sender, subject, attachments, body=""):
        self.mid, self.received, self.sender, self.subject = mid, received, sender, subject
        self.attachments = attachments          # [(파일명, 크기, save(path))]
        self.body = body                        # 대용량 링크 감지용으로만(저장하지 않음)


def eml_source(folder):
    for fn in sorted(os.listdir(folder)):
        if not fn.lower().endswith(".eml"):
            continue
        with open(os.path.join(folder, fn), "rb") as f:
            msg = email.message_from_binary_file(f, policy=email.policy.default)
        mid = msg.get("Message-ID") or hashlib.md5(open(os.path.join(folder, fn), "rb").read()).hexdigest()
        try:
            rec = parsedate_to_datetime(msg.get("Date")).replace(tzinfo=None)
        except Exception:
            rec = dt.datetime.fromtimestamp(os.path.getmtime(os.path.join(folder, fn)))
        atts = []
        for part in msg.iter_attachments():
            name = part.get_filename()
            if not name:
                continue
            data = part.get_payload(decode=True) or b""
            atts.append((name, len(data), (lambda d: (lambda p: open(p, "wb").write(d)))(data)))
        body = msg.get_body(preferencelist=("plain", "html"))
        yield Mail(mid, rec, parseaddr(msg.get("From", ""))[1], str(msg.get("Subject", "")), atts,
                   body.get_content() if body else "")


SKIP_DEFAULT = {3, 4, 5, 16, 23, 25}      # 지운 편지함·보낼 편지함·보낸 편지함·임시 보관함·정크·RSS
SKIP_NAMES = re.compile(r"지운|삭제|휴지통|정크|스팸|보낸|임시\s*보관|보관할\s*편지|Deleted|Junk|Spam|Sent|Draft|Outbox|"
                        r"Conversation|Sync|RSS|Archive Settings", re.I)


def outlook_folders(store, cfg):
    """사서함 안의 메일 폴더 전부(하위 폴더 포함). 지운 편지함·보낸 편지함·정크·일정·연락처 등은 뺀다.
    설정 outlook_folder: "auto"(기본, 전부) | 폴더 이름 또는 '받은 편지함/수탁사' 같은 경로(문자열·목록) | 없으면 auto."""
    skip_ids = set()
    for k in SKIP_DEFAULT:
        try:
            skip_ids.add(store.GetDefaultFolder(k).EntryID)
        except Exception:
            pass
    excl = [x.lower() for x in cfg.get("exclude_mail_folders") or []]
    out = []

    def walk(f, path):
        try:
            if f.EntryID in skip_ids or SKIP_NAMES.search(f.Name) or f.DefaultItemType != 0:   # 0 = 메일 폴더
                return
        except Exception:
            return
        p = f"{path}/{f.Name}" if path else f.Name
        if any(e in p.lower() for e in excl):
            return
        out.append((p, f))
        try:
            for sub in f.Folders:
                walk(sub, p)
        except Exception:
            pass

    root = store.GetRootFolder()
    for f in root.Folders:
        walk(f, "")
    want = cfg.get("outlook_folder", "auto")
    if want and want not in ("auto", "*", "전체"):
        names = [want] if isinstance(want, str) else list(want)
        hit = [(p, f) for p, f in out if any(n == f.Name or p.endswith(n) for n in names)]
        if hit:
            return hit
        print(f"  [안내] 설정 outlook_folder {names}가 없어 메일 폴더 전부를 봄")
    return out


def relevant(cfg, R, sender, subject, att_names):
    """SPC 자료 메일인가(폴더 전체를 볼 때 거르는 기준): 첨부가 있고, 보낸이가 수탁사·은행 도메인이거나
    제목·첨부 이름에 SPC 이름 또는 자료 종류(자금판·거래내역 등)가 있음."""
    if not att_names:
        return False
    dom = sender.split("@")[-1].lower()
    doms = {d.lower() for d in (cfg.get("domains") or {})} | set(BANK_DOMAINS) | set(PREPARER_DOMAINS) \
        | {d.lower() for d in (cfg.get("preparer_domains") or {})}
    if any(dom == d or dom.endswith("." + d) for d in doms):
        return True
    text = " ".join([subject] + list(att_names))
    if R.spc(subject, *att_names):
        return True
    return any(re.search(rx, text) for _, _, rx in KINDS)


def _since_filter(days):
    if not days:
        return None
    d = dt.datetime.now() - dt.timedelta(days=days)
    return f"[ReceivedTime] >= '{d:%m/%d/%Y} 00:00 AM'"


def _items(folder, days):
    items = folder.Items
    flt = _since_filter(days)
    if flt:
        try:
            items = items.Restrict(flt)
        except Exception:
            pass
    try:
        items.Sort("[ReceivedTime]", False)
    except Exception:
        pass
    return items


def _addr(m):
    addr = m.SenderEmailAddress or ""
    if m.SenderEmailType == "EX":
        try:
            addr = m.Sender.GetExchangeUser().PrimarySmtpAddress
        except Exception:
            pass
    return addr


def outlook_source(cfg, days=None):
    import win32com.client
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    try:
        ns.SendAndReceive(False)
    except Exception:
        pass
    R = Router(cfg)
    auto = cfg.get("outlook_folder", "auto") in ("auto", "*", "전체", None, "")
    seen = set()                                # 같은 메일이 여러 폴더에 복사돼 있으면 한 번만(인터넷 Message-ID)
    stores = [s for s in ns.Stores if cfg.get("outlook_store", "") in s.DisplayName]
    for store in stores:
        for path, folder in outlook_folders(store, cfg):
            for m in _items(folder, days):
                try:
                    if m.Class != 43:
                        continue
                    names = [a.FileName for a in m.Attachments]
                except Exception:
                    continue
                addr, subj = _addr(m), (m.Subject or "").strip()
                if auto and not relevant(cfg, R, addr, subj, names):
                    continue
                try:
                    imid = m.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
                except Exception:
                    imid = None
                if imid and imid in seen:
                    continue
                seen.add(imid)
                rt = m.ReceivedTime
                rec = dt.datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute)
                atts = []
                for a in m.Attachments:
                    try:
                        atts.append((a.FileName, a.Size, (lambda att: (lambda p: att.SaveAsFile(p)))(a)))
                    except Exception:
                        continue
                yield Mail(imid or m.EntryID, rec, addr, subj, atts, m.Body or "")


def survey(cfg, days=180):
    """편지함 조사(저장 없음): 메일 폴더마다 최근 N일 메일 수·첨부 있는 메일·SPC 자료로 보이는 메일 수와
    그 보낸이 도메인 상위. 결과는 화면과 <dest_root>/_수집작업/_편지함조사.md(로컬 전용)."""
    import win32com.client
    from collections import Counter
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    R = Router(cfg)
    rows, dom_all = [], Counter()
    for store in ns.Stores:
        for path, folder in outlook_folders(store, dict(cfg, outlook_folder="auto")):
            n = na = nr = 0
            doms = Counter()
            for m in _items(folder, days):
                try:
                    if m.Class != 43:
                        continue
                    names = [a.FileName for a in m.Attachments]
                    addr = _addr(m)
                except Exception:
                    continue
                n += 1
                na += bool(names)
                if relevant(cfg, R, addr, (m.Subject or ""), names):
                    nr += 1
                    doms[addr.split("@")[-1].lower()] += 1
            if n:
                rows.append((store.DisplayName, path, n, na, nr, doms.most_common(4)))
                dom_all.update(doms)
    rows.sort(key=lambda r: -r[4])
    lines = [f"# 편지함 조사 ({dt.datetime.now():%Y-%m-%d %H:%M}, 최근 {days}일)", "",
             "| 사서함 | 폴더 | 메일 | 첨부 있음 | SPC 자료로 보임 | 주요 보낸 도메인 |", "|---|---|---|---|---|---|"]
    for st, p, n, na, nr, d in rows:
        lines.append(f"| {st} | {p} | {n} | {na} | {nr} | {', '.join(f'{k} {v}' for k, v in d)} |")
    known = {x.lower() for x in (cfg.get("domains") or {})} | set(BANK_DOMAINS)
    unknown = [(k, v) for k, v in dom_all.most_common(30) if not any(k == d or k.endswith("." + d) for d in known)]
    lines += ["", "## 설정 domains에 없는 보낸 도메인(SPC 자료로 보인 메일 기준)", ""]
    lines += [f"- {k}: {v}통" for k, v in unknown] or ["- 없음"]
    total = sum(r[4] for r in rows)
    lines += ["", f"SPC 자료로 보인 메일 {total}통이 {sum(1 for r in rows if r[4])}개 폴더에 있음. "
              "수집기는 기본(outlook_folder: auto)으로 이 폴더들을 모두 보므로 편지함·규칙을 새로 만들 필요 없음."]
    out = "\n".join(lines)
    print(out)
    work = os.path.join(cfg["dest_root"], "_수집작업")
    os.makedirs(work, exist_ok=True)
    open(os.path.join(work, "_편지함조사.md"), "w", encoding="utf-8").write(out)
    return rows


# ---------------------------------------------------------------- 처리
def route_spc(R, path, base, msg_spc, subject, peek=True, extra=""):
    """SPC 판별 순서: 파일명 → 신보 약칭(2099-21) → 파일 내용(엑셀·CSV) → 메일 제목 → 이름 꼴(새 SPC). (spc, 새SPC여부, 근거)"""
    for t in (base, extra):                     # 파일명의 신보 SPC 이름(신보2025제3차)이 먼저 - 제목 SPC·느슨한 별칭보다 우선(10-23: 여러 SPC 자료를 한 메일로 받을 때 한 폴더로 몰리던 문제)
        names = {STRIP.sub("", m.group(0)) for m in SINBO_NAME.finditer(t or "")}
        if len(names) == 1:
            nm = names.pop()
            known = next((s for a, s in R.alias if a == nm or STRIP.sub("", s) == nm), None)   # 이름이 똑같을 때만
            return (known, False, "파일명") if known else (nm, True, "이름꼴")
    spc = R.spc(base) or (R.spc(extra) if extra else None)
    if spc:
        return spc, False, "파일명"
    sid = sinbo_id(base, f"{subject} {extra} {base}") or sinbo_id(extra, f"{subject} {extra}")
    if sid:
        known = R.spc(sid)
        return (known, False, "신보약칭") if known else (sid, True, "신보약칭")
    txt = peek_text(path) if (peek and base.lower().endswith(SHEET_EXT)) else ""
    if txt:
        names = R.spc_all(txt)
        if len(names) > 1:                      # 여러 SPC가 든 총괄표·리스트는 내용으로 정하지 않음
            txt = ""
        else:
            spc = R.spc(txt)
            if spc:
                return spc, False, "내용"
    if msg_spc:
        return msg_spc, False, "제목"
    for t in (base, extra, txt, subject):
        spc = R.new_spc(t)
        if spc:
            return spc, True, "내용" if t is txt else "이름꼴"
    return None, False, ""


def safe_move(src, dst):
    """옮기기. 사용 중이라 못 옮기면 잠깐 뒤 다시, 그래도 안 되면 건너뜀(반쯤 옮겨진 사본은 지움) - 한 파일 때문에 전체가 멈추지 않게(10-27)"""
    import gc, time
    for k in range(3):
        try:
            shutil.move(src, dst)
            return True
        except PermissionError:
            gc.collect()
            time.sleep(1 + k)
    if os.path.exists(src) and os.path.exists(dst) and filecmp.cmp(src, dst, shallow=False):
        try:
            os.remove(dst)
        except OSError:
            pass
    print(f"    ! 사용 중이라 못 옮김(그대로 둠): {os.path.basename(src)}")
    return False


def reroute(cfg, dry=False):
    """_미분류/ 에 있는 파일을 지금 규칙(파일명·내용·이름꼴)으로 다시 판별해 제자리로 옮긴다. 메일 다시 읽지 않음."""
    dest = cfg["dest_root"]
    R = Router(cfg)
    root = os.path.join(dest, "_미분류")
    moved, left, queue, failed = 0, 0, [], 0
    senders = {}                                  # 파일 이름 → 보낸이(_수신로그.md에서), 은행별 폴더 정할 때 사용
    lp = os.path.join(dest, "_수신로그.md")
    if os.path.exists(lp):
        for line in open(lp, encoding="utf-8"):
            c = [x.strip() for x in line.strip().strip("|").split("|")]
            if len(c) >= 5 and "@" in c[2]:
                senders[c[4]] = c[2]
    if not os.path.isdir(root):
        print("_미분류 폴더 없음")
        return
    for d in sorted(os.listdir(root)):
        dpath = os.path.join(root, d)
        if not os.path.isdir(dpath):
            continue
        subject = d.split("_", 1)[1] if "_" in d else ""
        try:
            received = dt.datetime.strptime(d[:8], "%Y%m%d")
        except ValueError:
            received = dt.datetime.now()
        for cur, _, files in list(os.walk(dpath)):
            items = []
            for base in files:
                p = os.path.join(cur, base)
                if base.lower().endswith(".zip") and not dry:
                    kids = extract_zip(p, os.path.splitext(p)[0])
                    if kids and kids != [p]:          # 풀렸으면 안의 파일로 분류, zip 원본은 작업 폴더로
                        keep = uniq(os.path.join(dest, "_수집작업", "풀린zip", d, base))
                        os.makedirs(os.path.dirname(keep), exist_ok=True)
                        shutil.move(p, keep)
                        items += [(k, os.path.basename(k), os.path.splitext(base)[0]) for k in kids]
                        continue
                items.append((p, base, ""))
            for p, base, par in items:
                spc, new_spc, how = route_spc(R, p, base, R.spc(subject), subject, extra=par)
                snd = senders.get(safe(base, 45)) or senders.get(safe(par + ".zip", 45), "")
                tru = R.trustee(snd, spc) or (R.inbox_home(spc) or R.bank_folder(snd, spc) if spc else None)
                if not spc:
                    left += 1
                    continue
                kind, sub = R.kind(base, par, subject)
                ptexts = [SINBO_ID.sub(" ", t) if how == "신보약칭" else t for t in (base, par, subject)]
                per = R.closing(spc, R.period(ptexts, received))
                target = uniq(os.path.join(dest, tru, spc, sub, per, base))
                print(f"    {kind:5s} {per} → {os.path.relpath(target, dest)} ({how}{', 새 SPC' if new_spc else ''})")
                if not dry:
                    os.makedirs(os.path.dirname(target), exist_ok=True)
                    if not safe_move(p, target):
                        failed += 1
                        continue
                    if kind in QUEUE_KINDS:
                        queue.append({"수신": f"{received:%Y-%m-%d}", "수탁사": tru, "SPC": spc, "새SPC": new_spc,
                                      "결산월": per, "종류": kind, "파일": os.path.relpath(target, dest)})
                moved += 1
    if queue:
        qp = os.path.join(dest, "_새자료.json")
        old = json.load(open(qp, encoding="utf-8")) if os.path.exists(qp) else []
        json.dump(old + queue, open(qp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"\n{'[모의] ' if dry else ''}다시 분류 {moved}개, 여전히 미분류 {left}개, 사용 중이라 못 옮김 {failed}개")


def closing_of(per, months):
    """자료 기간 YY.MM(끝에 * 있으면 받은 날로 추정) → 업무배정표 결산월 중 같거나 다음 결산월. 추정(*)이면 가까운 과거 결산월."""
    m = re.match(r"(\d\d)\.(\d\d)(\*?)", per or "")
    if not m or not months:
        return per
    y, mo, guess = 2000 + int(m.group(1)), int(m.group(2)), bool(m.group(3))
    for k in range(13):
        d = -k if guess else k
        t = y * 12 + (mo - 1) + d
        if t % 12 + 1 in months:
            return f"{(t // 12) % 100:02d}.{t % 12 + 1:02d}"
    return per


def copy_to_spc(R, spc, per, kind, src):
    """허용·확실한 규칙이면 그 SPC 기간 폴더에 사본(덮어쓰지 않음, 같은 내용 있으면 건너뜀) → 넣은 경로 또는 None"""
    try:
        d = R.spc_place(spc, per, kind)
    except Exception:
        return None
    if not d:
        return None
    t = os.path.join(d, os.path.basename(src))
    if os.path.exists(t) and filecmp.cmp(src, t, shallow=False):
        return None
    for f in os.listdir(d):                           # 이름만 바꿔 저장한 같은 파일이 이미 있으면 건너뜀
        fp = os.path.join(d, f)
        if os.path.isfile(fp) and os.path.getsize(fp) == os.path.getsize(src) and filecmp.cmp(src, fp, shallow=False):
            return None
    t = uniq(t)
    shutil.copy2(src, t)
    return t


def promote(R, dest, now):
    """수집함 <수탁사>/<SPC>/<종류>/<YY.MM>/ 의 자료 중 허용 SPC는 그 기간 폴더가 생겼으면 사본 → 배치로그 행"""
    out = []
    if not R.cfg.get("place_in_spc", False) or not R.cfg.get("place_allow"):
        return out
    for t in os.listdir(dest):
        tp = os.path.join(dest, t)
        if t.startswith("_") or not os.path.isdir(tp):
            continue
        for spc in os.listdir(tp):
            if spc not in R.spc_home:
                continue
            sp = os.path.join(tp, spc)
            for d, _, fs in os.walk(sp):
                per = os.path.basename(d)
                if not re.fullmatch(r"\d\d\.\d\d", per):
                    continue
                for f in fs:
                    kind, _ = R.kind(f, "", "")
                    cp = copy_to_spc(R, spc, per, kind, os.path.join(d, f))
                    if cp:
                        out.append([now, "(수집함에서)", safe(f, 60), cp])
    return out


def run(cfg, source, dry=False, baseline=False, days=None, no_queue=False, recheck=False):
    dest = cfg["dest_root"]
    work = os.path.join(dest, "_수집작업")
    state_p = os.path.join(work, "_처리완료_메일.json")
    st = json.load(open(state_p, encoding="utf-8")) if os.path.exists(state_p) else {"processed": {}}
    done = st["processed"]
    R = Router(dict(cfg, place_in_spc=False) if no_queue else cfg)   # 과거 메일 가져오기는 SPC 폴더에 넣지 않음(이미 손으로 정리한 과거 기간)
    try:                                               # 관리 대상 = 최신 업무배정표의 SPC·펀드·리츠
        import scope as _scope
        SC = _scope.load(cfg)
    except Exception as e:
        SC = None
        print(f"  [안내] 업무배정표를 읽지 못해 전체를 수집함: {str(e)[:80]}")
    R.scope = SC
    if SC:
        print(f"  관리 대상 {len(SC)}곳(업무배정표 {os.path.basename(SC.path)})")
    out_of_scope = 0
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    logrows, queue, n_new = [], [], 0
    placed = []                                         # 원래 SPC 폴더에 넣은 파일 [처리, 수신일, 파일, 위치]
    since = dt.datetime.now() - dt.timedelta(days=days) if days else None
    errors = 0
    try:
        for m in source:
            try:
                if (m.mid in done and not (recheck and done[m.mid].get("범위밖", 1) and not done[m.mid].get("baseline"))) \
                        or (since and m.received < since):
                    continue                          # recheck: 업무배정표 밖이라 건너뛴 첨부가 있던(또는 기록 없는) 메일만 다시 봄 - 같은 파일은 '이미 있음'으로 건너뜀
                oos0 = out_of_scope
                n_new += 1
                if baseline:
                    done[m.mid] = {"date": f"{m.received:%Y-%m-%d %H:%M}", "baseline": True}
                    continue
                short = f"{m.received:%Y-%m-%d}_{safe(m.sender.split('@')[0], 12)}_{safe(m.subject, 30)}"
                print(f"\n[신규] {m.received:%Y-%m-%d %H:%M} | {m.sender} | {m.subject[:50]}")
                tmp = os.path.join(work, "첨부원본", short)
                if os.path.isdir(tmp):                # 다시 보는 메일(--recheck-scope): 새 폴더에 원래 이름으로 받아야 '이미 있음'이 맞음(10-24 '(1)' 중복 문제)
                    tmp = uniq(tmp)
                files = []
                for name, size, save in m.attachments:
                    if is_junk(name, size):
                        continue
                    p = uniq(os.path.join(tmp, safe(name, 110)))
                    if not dry:
                        os.makedirs(tmp, exist_ok=True)
                        try:
                            save(p)
                        except Exception as e:
                            print(f"    ! 저장 실패 {name}: {e}")
                            continue
                    files += (extract_zip(p, os.path.splitext(p)[0]) if (p.lower().endswith(".zip") and not dry) else [p])
                msg_spc = R.spc(m.subject)
                for p in files:
                    base = os.path.basename(p)
                    par = os.path.basename(os.path.dirname(p)) if os.path.dirname(p) != tmp else ""   # 풀린 zip 이름
                    spc, new_spc, how = route_spc(R, p, base, msg_spc, m.subject, peek=not dry, extra=par)
                    snd = original_sender(m.sender, m.body, cfg)
                    tru = R.trustee(snd, spc)
                    if new_spc and not tru:
                        tru = R.inbox_home(spc) or R.bank_folder(snd, spc)
                    kind, sub = R.kind(base, par, m.subject)
                    ptexts = [SINBO_ID.sub(" ", t) if how == "신보약칭" else t for t in (base, par, m.subject)]
                    per = R.closing(spc, R.period(ptexts, m.received))
                    ent = SC.match(spc) if (SC and spc) else None
                    if SC and spc and not ent:                # 업무배정표에 없는 회사 자료는 저장하지 않음
                        out_of_scope += 1
                        continue
                    if spc and tru:
                        target = os.path.join(dest, tru, spc, sub, per, base)
                        note = ("새 SPC(SPC 폴더 없음)" if new_spc else "") + (" 파일 내용으로 판별" if how == "내용" else "") \
                            + (" 신보 약칭으로 판별" if how == "신보약칭" else "")
                        note = note.strip()
                    else:                                   # 같은 메일 첨부는 한 폴더에 모아 둔다
                        target = os.path.join(dest, "_미분류", f"{m.received:%Y%m%d}_{safe(m.subject, 30)}", base)
                        note = "SPC 미판별" if not spc else "수탁사 미판별"
                    if not dry and os.path.exists(target) and os.path.exists(p) and filecmp.cmp(p, target, shallow=False):
                        print(f"    (이미 있음) {os.path.relpath(target, dest)}")
                        continue
                    target = uniq(target)
                    print(f"    {kind:5s} {per} → {os.path.relpath(target, dest)}" + (f"  ({note})" if note else ""))
                    if not dry:
                        os.makedirs(os.path.dirname(target), exist_ok=True)
                        shutil.copy2(p, target)
                        if spc and not new_spc:             # 허용된 SPC는 원래 SPC 폴더(그 기간 폴더가 있을 때)에도 사본
                            cp = copy_to_spc(R, spc, per, kind, target)
                            if cp:
                                placed.append([now, f"{m.received:%Y-%m-%d}", safe(base, 60), cp])
                                print(f"          + SPC 폴더: {os.path.relpath(cp, R.cfg['spc_root'])}")
                        if spc and kind in QUEUE_KINDS:
                            queue.append({"수신": f"{m.received:%Y-%m-%d %H:%M}", "수탁사": tru, "SPC": spc, "새SPC": new_spc,
                                          "결산월": closing_of(per, ent["결산월"]) if ent else per, "종류": kind,
                                          "파일": os.path.relpath(target, dest)})
                    logrows.append([now, f"{m.received:%Y-%m-%d}", safe(m.sender, 24), safe(m.subject, 34),
                                    safe(base, 45), kind, f"{tru or '?'}/{spc or '?'}/{per}", note])
                big = re.search(r"대용량\s*파일\s*첨부\s*\d+개\s*\(([\d.]+ ?[MG]B)\)", m.body or "")
                if big:
                    logrows.append([now, f"{m.received:%Y-%m-%d}", safe(m.sender, 24), safe(m.subject, 34),
                                    f"(대용량 링크 {big.group(1)})", "**수동**", "", "링크 만료 전 직접 받아 _미분류에 넣기"])
                if not dry:
                    done[m.mid] = {"date": f"{m.received:%Y-%m-%d %H:%M}", "subject": safe(m.subject, 60), "files": len(files),
                                   "범위밖": out_of_scope - oos0}
                    if len(done) % 20 == 0:          # 작업 스케줄러 시간 제한으로 끊겨도 처리한 메일은 남게
                        os.makedirs(work, exist_ok=True)
                        json.dump(st, open(state_p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

            except Exception as e:                    # 메일 하나가 실패해도 나머지는 계속
                errors += 1
                print(f"    ! 처리 실패(다음 실행 때 다시 시도): {e}")
    finally:
        if not dry:                                   # 중간에 멈춰도 처리한 메일은 기록
            os.makedirs(work, exist_ok=True)
            json.dump(st, open(state_p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    if not dry:
        os.makedirs(work, exist_ok=True)
        json.dump(st, open(state_p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        if logrows:
            lp = os.path.join(dest, "_수신로그.md")
            new = not os.path.exists(lp)
            with open(lp, "a", encoding="utf-8") as f:
                if new:
                    f.write("# SPC 자료 메일 수신·분류 로그\n\n`ingest_spc.py`가 기록합니다. 분류가 틀리면 파일을 옮기고 비고에 적어 주세요"
                            " - 반복되면 설정(aliases·domains)에 추가합니다.\n\n"
                            "| 처리 | 수신일 | 보낸이 | 제목 | 파일 | 종류 | 수탁사/SPC/결산월 | 비고 |\n|---|---|---|---|---|---|---|---|\n")
                for r in logrows:
                    f.write("| " + " | ".join(r) + " |\n")
        if not baseline and not no_queue:            # 예전에 받아 수집함에만 있던 자료: 이제 그 기간 폴더가 생겼으면 사본
            placed += promote(R, dest, now)
        if placed:
            pp = os.path.join(dest, "_SPC폴더_배치로그.md")
            new = not os.path.exists(pp)
            with open(pp, "a", encoding="utf-8") as f:
                if new:
                    f.write("# 원래 SPC 폴더에 넣은 파일\n\n수집기가 사용자 SPC 폴더(기간 폴더·회사제시)에 추가한 파일입니다. "
                            "덮어쓴 파일은 없습니다(같은 이름이면 (1)을 붙임). 잘못 들어갔으면 이 목록으로 찾아 옮기세요.\n\n"
                            "| 처리 | 수신일 | 파일 | 위치 |\n|---|---|---|---|\n")
                for r in placed:
                    f.write("| " + " | ".join(r) + " |\n")
        if queue and not no_queue:                   # 생성 파이프라인 대기열(새 자금판·거래내역)
            qp = os.path.join(dest, "_새자료.json")
            old = json.load(open(qp, encoding="utf-8")) if os.path.exists(qp) else []
            json.dump(old + queue, open(qp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    if out_of_scope:
        print(f"업무배정표에 없는 회사 자료 {out_of_scope}개는 저장하지 않음")
    print(f"\n{'[모의] ' if dry else ''}신규 메일 {n_new}건" + (" - 처리완료 표시만" if baseline else "")
          + (f", 실패 {errors}건" if errors else ""))
    return logrows, queue


def main():
    a = sys.argv[1:]
    opt = lambda k: a[a.index(k) + 1] if k in a else None
    cfg = json.load(open(opt("--config") or os.path.join(HERE, "spc_mail_config.json"), encoding="utf-8"))
    load_zip_passwords(cfg)
    if "--reroute" in a:
        return reroute(cfg, dry="--dry-run" in a)
    days = int(opt("--days")) if opt("--days") else None
    if "--survey" in a:
        return survey(cfg, days or 180)
    src = eml_source(opt("--eml")) if opt("--eml") else outlook_source(cfg, days)
    run(cfg, src, dry="--dry-run" in a, baseline="--baseline" in a, days=days, no_queue="--no-queue" in a,
        recheck="--recheck-scope" in a)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    try:
        main()
    except Exception:
        traceback.print_exc()
        sys.exit(1)
