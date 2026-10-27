# -*- coding: utf-8 -*-
"""리뷰어 메일 모으기(회사 PC 전용, pc_agent '리뷰수집' 작업) - 10-27 사용자 요청.

설정 spc_mail_config.json "reviewers": ["이름", ...] (보낸 사람 표시 이름에 들어 있으면 리뷰 메일).
Outlook 메일 폴더 전부(보낸·지운 편지함 제외)에서 그 사람들이 보낸 메일을 찾아
 수집함\\_리뷰메일\\<SPC 또는 _SPC미정>\\<받은날>_<리뷰어>_<제목>.txt  (본문에서 인용한 원문은 뺌)
 + 첨부(리뷰 시트 등)는 같은 이름 폴더에, 목록은 수집함\\_리뷰메일\\_목록.json
메일은 PC 밖으로 보내지 않는다(분석 결과만 SPC-n으로 가려서 저장소에).
  python review_mail.py [--days N]
"""
import os, re, sys, json, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import ingest_spc as I
import scope as _scope

QUOTE = re.compile(r"^\s*(-{2,}\s*(Original Message|원본 메시지|원래 메시지)\s*-{2,}|From:\s|보낸 사람:\s|"
                   r"On .+ wrote:|20\d\d[-./] ?\d+[-./] ?\d+.*작성:)", re.I | re.M)


def own_text(body):
    """회신 메일에서 리뷰어가 새로 쓴 부분만(아래 인용한 원문 앞까지)"""
    body = (body or "").replace("\r\n", "\n")
    m = QUOTE.search(body)
    t = body[:m.start()] if m else body
    return re.sub(r"\n{3,}", "\n\n", t).strip()


def main():
    a = sys.argv[1:]
    days = int(a[a.index("--days") + 1]) if "--days" in a else None
    cfg = json.load(open(os.path.join(HERE, "spc_mail_config.json"), encoding="utf-8"))
    who = cfg.get("reviewers") or []
    if not who:
        print("설정 reviewers 없음")
        return
    dest = os.path.join(os.path.normpath(cfg["dest_root"]), "_리뷰메일")
    os.makedirs(dest, exist_ok=True)
    R = I.Router(cfg)
    sc = _scope.load(cfg)
    import win32com.client
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    seen, rows = set(), []
    stores = [s for s in ns.Stores if cfg.get("outlook_store", "") in s.DisplayName]
    for store in stores:
        for path, folder in I.outlook_folders(store, dict(cfg, outlook_folder="auto")):
            for m in I._items(folder, days):
                try:
                    if m.Class != 43:
                        continue
                    name = m.SenderName or ""
                except Exception:
                    continue
                rv = next((w for w in who if w in name), None)
                if not rv:
                    continue
                try:
                    imid = m.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
                except Exception:
                    imid = m.EntryID
                if imid in seen:
                    continue
                seen.add(imid)
                subj = (m.Subject or "").strip()
                body = m.Body or ""
                att = []
                try:
                    att = [x.FileName for x in m.Attachments]
                except Exception:
                    pass
                spc = R.spc(subj, *att) or R.spc(own_text(body)[:3000]) or R.spc(body[:20000])
                insc = bool(sc and spc and sc.match(spc))
                rt = m.ReceivedTime
                stamp = f"{rt.year:04d}{rt.month:02d}{rt.day:02d}_{rt.hour:02d}{rt.minute:02d}"
                folder_name = I.safe(spc, 60) if spc else "_SPC미정"
                base = f"{stamp}_{rv}_{I.safe(subj, 50)}"
                d = os.path.join(dest, folder_name)
                os.makedirs(d, exist_ok=True)
                p = os.path.join(d, base + ".txt")
                own = own_text(body)
                with open(p, "w", encoding="utf-8") as f:
                    f.write(f"보낸이: {rv}\n받은날: {rt:%Y-%m-%d %H:%M}\n제목: {subj}\n폴더: {path}\n첨부: {', '.join(att)}\n\n")
                    f.write("== 새로 쓴 부분 ==\n" + own + "\n\n== 원문 전체(인용 포함) ==\n" + body.replace("\r\n", "\n"))
                saved = []
                for x in getattr(m, "Attachments", []):
                    try:
                        fn = x.FileName
                        if I.is_junk(fn, getattr(x, "Size", None)):
                            continue
                        ad = os.path.join(d, base)
                        os.makedirs(ad, exist_ok=True)
                        x.SaveAsFile(os.path.join(ad, I.safe(fn, 80)))
                        saved.append(fn)
                    except Exception:
                        continue
                rows.append({"리뷰어": rv, "받은날": f"{rt:%Y-%m-%d %H:%M}", "제목": subj, "SPC": spc, "배정표대상": insc,
                             "파일": os.path.relpath(p, dest), "첨부": saved, "새로쓴글자수": len(own)})
    rows.sort(key=lambda r: r["받은날"])
    json.dump(rows, open(os.path.join(dest, "_목록.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    from collections import Counter
    print(f"리뷰 메일 {len(rows)}통: " + ", ".join(f"{k} {v}" for k, v in Counter(r['리뷰어'] for r in rows).items()))
    print(f"SPC 정해짐 {sum(1 for r in rows if r['SPC'])}통 ({len({r['SPC'] for r in rows if r['SPC']})}곳), 미정 {sum(1 for r in rows if not r['SPC'])}통")


if __name__ == "__main__":
    main()
