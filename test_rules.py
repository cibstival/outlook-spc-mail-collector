# -*- coding: utf-8 -*-
"""SPC 폴더 넣기 규칙 시험(가상 SPC만, 고객 자료 없음).

사용자가 예전에 메일 첨부를 저장한 위치에서 규칙을 배우고, 허용한 SPC에만, 그 기간 폴더가 이미 있을 때만 사본을 둔다.
- 기수 틀 AT/[AT]x_{N}기(YYYY.MM)/회사제시, 날짜 틀 T/[T]x_YYYY.MM, 이름 바꿔 저장한 파일도 내용으로 짝
- 허용 안 한 SPC·확실하지 않은 규칙·폴더 없는 기간은 넣지 않음(수집함에만)
- 폴더가 나중에 생기면 다음 실행 때 사본(promote)
"""
import os, sys, json, shutil, tempfile, subprocess, datetime as dt, email.message, email.utils

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import spc_rules as R
import ingest_spc as I


def w(p, data):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    open(p, "wb").write(data)


def mail(folder, n, frm, subj, atts, date):
    m = email.message.EmailMessage()
    m["From"], m["To"], m["Subject"] = frm, "me@example.com", subj
    m["Date"], m["Message-ID"] = email.utils.format_datetime(date), f"<{n}@test>"
    m.set_content("본문")
    for fn, data in atts:
        m.add_attachment(data, maintype="application", subtype="octet-stream", filename=fn)
    open(os.path.join(folder, f"{n}.eml"), "wb").write(bytes(m))


def main():
    ok = True
    def check(c, msg):
        nonlocal ok
        print(("  OK  " if c else "  실패 ") + msg)
        ok &= bool(c)

    tmp = tempfile.mkdtemp()
    cwd = os.getcwd()
    try:
        os.chdir(tmp)
        S = "SPC/가나은행/한빛제일차"
        for d, yy, n in (("2025.06", "25.06", 7), ("2025.12", "25.12", 7), ("2026.06", "26.06", 8)):
            cf = ("cf" + d).encode() * 600
            w(f"{S}/AT/[AT]한빛_{n}기({d})/회사제시/한빛_자금판_{d}.xlsx", cf)
            w(f"수집함/가나은행/한빛제일차/자금판 및 거래내역/{yy}/한빛_자금판_{d}.xlsx", cf)
            tx = ("tax" + d).encode() * 700
            w(f"{S}/T/[T]한빛_{d}/계산서합계.pdf", tx)
            w(f"수집함/가나은행/한빛제일차/세금계산서/{yy}/세금계산서합계표({d}).pdf", tx)
        w("SPC/다라증권/두리제이차/2026.06/A.두리(2026.06).xlsx", b"x" * 3000)   # 허용 안 할 SPC
        data = R.learn("SPC", "수집함", lambda f: I.Router.kind(f, "", "")[0], os.path.join(tmp, "rules.json"), log=lambda *_: None)
        r = data["규칙"]
        check(r["가나은행/한빛제일차|자금판"]["틀"] == "AT/[AT]한빛_{N}기({YYYY.MM+0})/회사제시" and r["가나은행/한빛제일차|자금판"]["확실"], "자금판 규칙: 기수 틀")
        check(r["가나은행/한빛제일차|세금계산서"]["틀"] == "T/[T]한빛_{YYYY.MM+0}", "세금계산서 규칙: 이름 바꿔 저장한 것도 내용으로 짝")
        sd = os.path.join("SPC", "가나은행", "한빛제일차")
        check(R.target_dir(sd, r["가나은행/한빛제일차|자금판"]["틀"], "26.06").endswith("8기(2026.06)/회사제시".replace("/", os.sep)), "26.06 → 8기(2026.06)/회사제시")
        check(R.target_dir(sd, r["가나은행/한빛제일차|자금판"]["틀"], "26.12") is None, "26.12 폴더 없음 → None")

        eml = os.path.join(tmp, "eml"); os.makedirs(eml)
        mail(eml, 1, "x@preparer.example.kr", "한빛제일차 12월 자금판", [("한빛_자금판_2026.12.xlsx", b"new" * 900)], dt.datetime(2027, 1, 10, 10))
        mail(eml, 2, "x@preparer.example.kr", "두리제이차 A", [("A.두리(2026.12).xlsx", b"a" * 3000)], dt.datetime(2027, 1, 10, 11))
        cfgp = os.path.join(tmp, "cfg.json")
        json.dump({"dest_root": os.path.join(tmp, "수집함"), "spc_root": os.path.join(tmp, "SPC"), "place_in_spc": True,
                   "place_allow": ["가나은행/한빛제일차"], "place_rules": os.path.join(tmp, "rules.json")}, open(cfgp, "w"))
        run = lambda: subprocess.run([sys.executable, os.path.join(HERE, "ingest_spc.py"), "--config", cfgp, "--eml", eml],
                                     capture_output=True, check=True)
        run()
        check(os.path.exists("수집함/가나은행/한빛제일차/자금판 및 거래내역/26.12/한빛_자금판_2026.12.xlsx"), "수집함에는 항상 저장")
        check(not os.path.exists(f"{S}/AT/[AT]한빛_8기(2026.12)"), "기간 폴더가 없으면 SPC 폴더에 만들지 않음")
        check(not any("A.두리(2026.12)" in f for _, _, fs in os.walk("SPC/다라증권") for f in fs), "허용 안 한 SPC는 넣지 않음")
        os.makedirs(f"{S}/AT/[AT]한빛_8기(2026.12)/회사제시")                      # 사용자가 그 기간 폴더를 만듦
        mail(eml, 3, "x@preparer.example.kr", "기타", [("메모.txt", b"m" * 10)], dt.datetime(2027, 1, 11, 10))
        run()
        check(os.path.exists(f"{S}/AT/[AT]한빛_8기(2026.12)/회사제시/한빛_자금판_2026.12.xlsx"), "폴더가 생기면 다음 실행 때 사본(promote)")
        run()
        n = len(os.listdir(f"{S}/AT/[AT]한빛_8기(2026.12)/회사제시"))
        check(n == 1, "다시 실행해도 사본이 늘지 않음")
        check(os.path.exists("수집함/_SPC폴더_배치로그.md"), "배치로그 기록")
    finally:
        os.chdir(cwd)
        shutil.rmtree(tmp, ignore_errors=True)
    print("결과:", "통과" if ok else "실패")
    return ok


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(0 if main() else 1)
