# -*- coding: utf-8 -*-
"""수집함 중복 사본 정리(10-24): 'name(1).ext'가 같은 폴더의 'name.ext'와 내용이 똑같으면 수집함\\_중복\\<원래 경로>로 옮김(지우지 않음).
대기열 _새자료.json에서 그 사본 항목은 뺌. 옮긴 목록은 수집함\\_중복_정리.md
  python dedupe_inbox.py --dest <수집함> [--apply]"""
import os, re, sys, json, hashlib, shutil, argparse, datetime as dt

SUF = re.compile(r"^(.*)\((\d+)\)(\.[^.\\/]+)$")


def md5(p):
    h = hashlib.md5()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dest", required=True)
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()
    dest = a.dest
    moves = []
    for t in sorted(os.listdir(dest)):
        tp = os.path.join(dest, t)
        if t.startswith("_") or not os.path.isdir(tp):
            continue
        for cur, _, files in os.walk(tp):
            names = set(files)
            for f in files:
                m = SUF.match(f)
                if not m:
                    continue
                base = m.group(1) + m.group(3)
                if base not in names:
                    continue
                p, b = os.path.join(cur, f), os.path.join(cur, base)
                try:
                    if os.path.getsize(p) == os.path.getsize(b) and md5(p) == md5(b):
                        moves.append(p)
                except OSError:
                    pass
    print(f"똑같은 사본 {len(moves)}개")
    if not a.apply:
        return
    qp = os.path.join(dest, "_새자료.json")
    q = json.load(open(qp, encoding="utf-8")) if os.path.exists(qp) else []
    rel = lambda p: os.path.relpath(p, dest).replace("/", "\\")
    gone = {rel(p) for p in moves}
    n0 = len(q)
    q = [e for e in q if e.get("파일") not in gone]
    log = [f"# 중복 사본 정리 {dt.datetime.now():%Y-%m-%d %H:%M}", ""]
    for p in moves:
        to = os.path.join(dest, "_중복", os.path.relpath(p, dest))
        os.makedirs(os.path.dirname(to), exist_ok=True)
        shutil.move(p, to)
        log.append(f"- {rel(p)}")
    json.dump(q, open(qp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    open(os.path.join(dest, "_중복_정리.md"), "a", encoding="utf-8").write("\n".join(log) + "\n")
    print(f"옮김 {len(moves)}개, 대기열 {n0} → {len(q)}")


if __name__ == "__main__":
    main()
