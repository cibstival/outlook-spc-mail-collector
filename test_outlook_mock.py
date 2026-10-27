# -*- coding: utf-8 -*-
"""Outlook COM 없이 폴더 탐색·SPC 메일 거르기 시험(가짜 사서함)."""
import sys, os, types, datetime as dt, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ingest_spc as I

class Att:
    def __init__(s, n): s.FileName, s.Size = n, 10
    def SaveAsFile(s, p): open(p, "wb").write(b"x")
class Msg:
    Class = 43
    def __init__(s, frm, subj, atts, imid):
        s.SenderEmailAddress, s.SenderEmailType, s.Subject = frm, "SMTP", subj
        s.Attachments = [Att(a) for a in atts]; s.ReceivedTime = dt.datetime(2026, 10, 1, 9, 0); s.Body = ""
        s.EntryID = imid + "-e"; s.PropertyAccessor = types.SimpleNamespace(GetProperty=lambda k, i=imid: i)
class Items(list):
    def Restrict(s, f): return s
    def Sort(s, *a): pass
class Folder:
    def __init__(s, name, msgs=(), subs=(), kind=0, eid=None):
        s.Name, s.Items, s.Folders, s.DefaultItemType, s.EntryID = name, Items(msgs), list(subs), kind, eid or name
class Store:
    DisplayName = "me@firm.kr"
    def __init__(s, root, defaults): s._root, s._d = root, defaults
    def GetRootFolder(s): return s._root
    def GetDefaultFolder(s, k): return s._d[k]

spc = Msg("ops@gana-sec.co.kr", "[가나다제일차] 9월 자금판", ["CF_260930.xlsx"], "m1")
dup = Msg("ops@gana-sec.co.kr", "[가나다제일차] 9월 자금판", ["CF_260930.xlsx"], "m1")   # 다른 폴더의 같은 메일
bank = Msg("x@ibk.co.kr", "거래내역", ["2026-9.xls"], "m2")
noise = Msg("friend@gmail.com", "점심", [], "m3")
noatt = Msg("ops@gana-sec.co.kr", "확인 부탁", [], "m4")
sent = Folder("보낸 편지함", [Msg("me@firm.kr", "자금판 회신", ["a.xlsx"], "m5")], kind=0, eid="SENT")
cal = Folder("일정", kind=1)
inbox = Folder("받은 편지함", [noise, noatt], [Folder("수탁사", [spc]), Folder("은행", [bank]), Folder("개인", [dup])])
root = Folder("root", subs=[inbox, sent, cal, Folder("지운 편지함", [spc], eid="DEL")])
store = Store(root, {3: Folder("x", eid="DEL"), 4: Folder("x", eid="OUT"), 5: sent, 16: Folder("x", eid="DR"), 23: Folder("x", eid="J"), 25: Folder("x", eid="R")})

tmp = tempfile.mkdtemp()
os.makedirs(os.path.join(tmp, "SPC", "가나증권", "가나다제일차유동화전문(유)"))
cfg = {"dest_root": os.path.join(tmp, "수집함"), "spc_root": os.path.join(tmp, "SPC"),
       "domains": {"gana-sec.co.kr": "가나증권"}, "outlook_folder": "auto"}
paths = [p for p, f in I.outlook_folders(store, cfg)]
print("폴더:", paths)
fake = types.SimpleNamespace(Dispatch=lambda n: types.SimpleNamespace(GetNamespace=lambda x: types.SimpleNamespace(
    Stores=[store], SendAndReceive=lambda b: None)))
sys.modules["win32com"] = types.SimpleNamespace(client=fake); sys.modules["win32com.client"] = fake
got = [(m.sender, m.subject) for m in I.outlook_source(cfg)]
print("수집:", got)
ok = (paths == ["받은 편지함", "받은 편지함/수탁사", "받은 편지함/은행", "받은 편지함/개인"]
      and got == [("ops@gana-sec.co.kr", "[가나다제일차] 9월 자금판"), ("x@ibk.co.kr", "거래내역")])
cfg2 = dict(cfg, outlook_folder=["수탁사"])
ok &= [p for p, f in I.outlook_folders(store, cfg2)] == ["받은 편지함/수탁사"]
I.survey(cfg, 30)
print("결과:", "통과" if ok else "실패")
sys.exit(0 if ok else 1)
