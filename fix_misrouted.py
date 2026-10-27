# -*- coding: utf-8 -*-
"""수집함에서 다른 SPC 폴더에 잘못 들어간 파일 바로잡기(10-23).

파일 이름에 '신보YYYY제N차'가 있는데 그 폴더의 SPC와 다르면:
- 그 SPC가 업무배정표 대상이면 수집함\\<그 SPC가 있는 수탁사>\\<그 SPC>\\<같은 하위 경로>로 옮김
- 대상이 아니면 수집함\\_범위밖\\<원래 경로>로 옮김(지우지 않음)
- 대기열 _새자료.json의 경로·SPC도 같이 고침. 옮긴 목록은 수집함\\_잘못분류_바로잡음.md

  python fix_misrouted.py --dest <수집함> [--assign <업무배정표 폴더>] [--apply]
"""
import os, re, sys, json, shutil, argparse, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import scope as _scope

SINBO_NAME = re.compile(r"신보\s*\d{4}\s*제\s*\d+\s*차")
STRIP = re.compile(r"유동화\s*전문|유한\s*회사|주식\s*회사|\(\s*(주|유)\s*\)|㈜|\s")
key = lambda s: STRIP.sub("", s or "")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dest", required=True)
    ap.add_argument("--assign")
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()
    dest = a.dest
    SC = _scope.load({"assignment_dir": a.assign}) if a.assign else None
    folders, pairs = {}, []                        # 키 → (수탁사, SPC 폴더 이름) / 같은 SPC가 여러 수탁사 폴더에 있어도 모두 봄
    for t in sorted(os.listdir(dest)):
        tp = os.path.join(dest, t)
        if t.startswith("_") or not os.path.isdir(tp):
            continue
        for s in os.listdir(tp):
            if os.path.isdir(os.path.join(tp, s)):
                folders.setdefault(key(s), (t, s))
                pairs.append((key(s), t, s))
    moves = []
    for k, t, s in pairs:
        base = os.path.join(dest, t, s)
        for cur, _, files in os.walk(base):
            for f in files:
                names = {key(m.group(0)) for m in SINBO_NAME.finditer(f)}
                if not names or any(n == k or k.startswith(n) and not k[len(n):len(n) + 1].isdigit() for n in names):
                    continue
                if len(names) > 1:
                    continue                       # 여러 SPC 이름 - 사람이 판단
                n = names.pop()
                rel = os.path.relpath(os.path.join(cur, f), base)
                hit = next((v for kk, v in folders.items() if kk == n or (kk.startswith(n) and not kk[len(n):len(n) + 1].isdigit())), None)
                inscope = (SC.match(n) is not None) if SC else True
                if inscope and hit:
                    to = os.path.join(dest, hit[0], hit[1], rel)
                elif inscope:
                    to = os.path.join(dest, t, n, rel)
                else:
                    to = os.path.join(dest, "_범위밖", t, s, rel)
                moves.append((os.path.join(cur, f), to, n, inscope))
    print(f"잘못 분류 {len(moves)}개: 대상 SPC로 {sum(1 for m in moves if m[3])}개, 범위 밖 {sum(1 for m in moves if not m[3])}개")
    if not a.apply:
        return
    qp = os.path.join(dest, "_새자료.json")
    q = json.load(open(qp, encoding="utf-8")) if os.path.exists(qp) else []
    rel = lambda p: os.path.relpath(p, dest).replace("/", "\\")
    qi = {e.get("파일"): e for e in q}
    log = [f"# 잘못 분류 바로잡음 {dt.datetime.now():%Y-%m-%d %H:%M}", ""]
    done = 0
    for src, to, n, inscope in moves:
        if os.path.exists(to):
            root, ext = os.path.splitext(to)
            to = f"{root}_이동{ext}"
        os.makedirs(os.path.dirname(to), exist_ok=True)
        shutil.move(src, to)
        done += 1
        e = qi.get(rel(src))
        if e is not None:
            if inscope:
                e["파일"] = rel(to)
                parts = rel(to).split("\\")
                e["수탁사"], e["SPC"] = parts[0], parts[1]
            else:
                q.remove(e)
        log.append(f"- {rel(src)} → {rel(to)}")
    json.dump(q, open(qp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    open(os.path.join(dest, "_잘못분류_바로잡음.md"), "a", encoding="utf-8").write("\n".join(log) + "\n")
    print(f"옮김 {done}개")


if __name__ == "__main__":
    main()
