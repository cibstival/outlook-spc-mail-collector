# -*- coding: utf-8 -*-
"""ingest_spc 시험 - 가상 SPC·가상 메일(.eml)로 분류 규칙을 확인한다(고객 자료 없음).

python test_ingest.py   → 임시 폴더에 SPC 루트·메일을 만들고 수집 후 기대값과 대조
"""
import os, sys, json, tempfile, zipfile, io, datetime as dt
from email.message import EmailMessage
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ingest_spc as I

CASES = [  # (보낸이, 제목, 첨부, 기대 수탁사, 기대 SPC, 기대 종류, 기대 결산월)
    ("ops@gana-sec.co.kr", "[가나다제일차] 6월 자금판 송부", ["(ABS) 가나다제일차유동화전문_CF_260630_v1.xlsx"],
     "가나증권", "가나다제일차유동화전문(유)", "자금판", "26.06"),
    ("ops@gana-sec.co.kr", "거래내역 송부의 건", ["가나다제일차유동화전문(유) 거래내역_1010-0000-0000(2026-01-01~2026-06-30).xls"],
     "가나증권", "가나다제일차유동화전문(유)", "거래내역", "26.06"),
    ("trust@raon-bank.com", "[라온스타제2차]_자금운용지시_12차(대출이자 지급)_송부의 건_20260305", ["지시서.pdf"],
     "라온은행", "라온스타제이차(주)", "운용지시서", "26.03"),
    ("trust@raon-bank.com", "자료 송부", ["라온스타2차_자금관리.zip"],          # zip 안: 자금관리 + 통장
     "라온은행", "라온스타제이차(주)", None, "26.06"),
    ("someone@unknown.com", "문의드립니다", ["견적서.pdf"], None, None, "기타", None),
    ("ops@gana-sec.co.kr", "잔액증명서", ["잔액증명서_가나다제일차_20260630.pdf", "image001.png"],
     "가나증권", "가나다제일차유동화전문(유)", "잔액증명", "26.06"),
]


def build(root):
    spc_root = os.path.join(root, "SPC")
    for t, s in (("가나증권", "가나다제일차유동화전문(유)"), ("라온은행", "라온스타제이차(주)")):
        os.makedirs(os.path.join(spc_root, t, s), exist_ok=True)
    mails = os.path.join(root, "mail")
    os.makedirs(mails)
    for k, (frm, subj, atts, *_) in enumerate(CASES):
        m = EmailMessage()
        m["From"], m["To"], m["Subject"] = frm, "me@firm.kr", subj
        m["Date"] = "Tue, 07 Jul 2026 10:00:00 +0900"
        m["Message-ID"] = f"<t{k}@test>"
        m.set_content("본문(저장되지 않아야 함)")
        for a in atts:
            if a.endswith(".zip"):
                b = io.BytesIO()
                with zipfile.ZipFile(b, "w") as z:
                    z.writestr("자금관리(라온스타제이차)_2606.xls", b"x")
                    z.writestr("라온스타2차_통장거래내역_260630.xlsx", b"x")
                data = b.getvalue()
            else:
                data = b"x" * 10
            m.add_attachment(data, maintype="application", subtype="octet-stream", filename=a)
        open(os.path.join(mails, f"{k}.eml"), "wb").write(bytes(m))
    cfg = {"dest_root": os.path.join(root, "수집함"), "spc_root": spc_root,
           "domains": {"gana-sec.co.kr": "가나증권", "raon-bank.com": "라온은행"},
           "aliases": {"라온스타제2차": "라온스타제이차(주)", "라온스타2차": "라온스타제이차(주)"}}
    return cfg, mails


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    root = tempfile.mkdtemp(prefix="ingest_test_")
    cfg, mails = build(root)
    rows, queue = I.run(cfg, I.eml_source(mails))
    bad = 0
    for k, (frm, subj, atts, tru, spc, kind, per) in enumerate(CASES):
        got = [r for r in rows if r[3] == I.safe(subj, 34)]
        for r in got:
            t, s, p = r[6].split("/")
            ok = (t == (tru or "?")) and (s == (spc or "?")) and (kind is None or r[5] == kind) and (per is None or p == per)
            bad += not ok
            print(f"  {'OK ' if ok else 'NG '} {r[4][:40]:40s} {r[5]:5s} {r[6]}")
    # 두 번째 실행은 아무것도 하지 않아야(중복 방지)
    rows2, _ = I.run(cfg, I.eml_source(mails))
    bad += bool(rows2)
    body_saved = any("본문" in open(os.path.join(dp, f), encoding="utf-8", errors="ignore").read()
                     for dp, _, fs in os.walk(cfg["dest_root"]) for f in fs if f.endswith((".md", ".json", ".txt")))
    bad += body_saved
    print(f"\n대기열 {len(queue)}건, 중복 재처리 {len(rows2)}건, 본문 저장 {'있음(오류)' if body_saved else '없음'}")
    print("결과:", "통과" if not bad else f"실패 {bad}건")
    sys.exit(1 if bad else 0)
