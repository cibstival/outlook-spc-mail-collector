# -*- coding: utf-8 -*-
"""회사 PC 요청 감시(5분마다, 창 없이): 원격 점검 세션이 사람 손 없이 PC에서 실행·시험할 수 있게.

원격 세션이 수집\\_요청\\<id>.json 을 써 두면 여기서 정해진 작업만 실행하고 결과를 수집함\\_요청결과\\<id>.txt 에 남긴다.
허용 작업(그 밖은 무시):
  {"작업": "자동실행"}  - auto_run.py (메일 수집 → 자동 결산 → 오후 과거 메일)
  {"작업": "시험"}      - 저장소의 test_*.py 전부 + 실행 환경 점검(파이썬·패키지·Outlook)
  {"작업": "중복정리"}  - dedupe_inbox.py --apply: 수집함의 똑같은 '(1)' 사본을 _중복으로 옮김(지우지 않음)
  {"작업": "리뷰수집"}  - review_mail.py: 리뷰어 메일을 SPC별로 수집함\\_리뷰메일에 모음
  {"작업": "전수시험", "대상": "SPC-2"}  - full_test.py(대상 생략하면 전부): 업무배정표 감사·검토 대상 전부를 정답 DSD와 대조(결과는 숫자·표 머리글만)
작업 스케줄러 등록은 auto_run.py 가 처음 돌 때 스스로 한다(사용자 계정, 로그인 중에만).
"""
import os, sys, json, glob, shutil, subprocess, datetime as dt, traceback

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
REQ = os.path.join(HERE, "_요청")
OUT = os.path.join(ROOT, "수집함", "_요청결과")
NO_WINDOW = 0x08000000 if os.name == "nt" else 0
PKGS = ["pywin32", "xlrd", "openpyxl", "pyzipper", "pymupdf"]


def py_cmd():
    uv = shutil.which("uv")
    if uv:
        return [uv, "run"] + sum((["--with", p] for p in PKGS), []) + ["python"]
    exe = sys.executable
    if exe.lower().endswith("pythonw.exe") and os.path.exists(exe[:-11] + "python.exe"):
        exe = exe[:-11] + "python.exe"
    return [exe]


def run(args, out, cwd, timeout=1800):
    out.write(f"\n$ {' '.join(os.path.basename(a) if os.path.isabs(a) else a for a in args)}\n")
    out.flush()
    env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8")
    try:
        r = subprocess.run(args, cwd=cwd, stdout=out, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                           env=env, timeout=timeout, creationflags=NO_WINDOW)
        out.write(f"[종료 코드 {r.returncode}]\n")
        return r.returncode
    except Exception as e:
        out.write(f"[실행 실패 {type(e).__name__}: {e}]\n")
        return -1
    finally:
        out.flush()


def selftest(out):
    py = py_cmd()
    out.write(f"python: {sys.version.split()[0]} {sys.executable}\nuv: {shutil.which('uv')}\n")
    run(py + ["-c", "import win32com.client as w; o=w.Dispatch('Outlook.Application'); "
              "print('Outlook 연결 OK', o.Session.Stores.Count, '개 사서함')"], out, HERE, 120)
    fails = 0
    tests = sorted(glob.glob(os.path.join(ROOT, "수집", "test_*.py")) +
                   glob.glob(os.path.join(ROOT, "표준화", "**", "test_*.py"), recursive=True))   # 저장소 코드 폴더만(수집함 사본 등 제외)
    for t in tests:
        fails += run(py + [t], out, os.path.dirname(t), 900) != 0
    out.write(f"\n시험 실패 {fails}건\n")


def keep_awake():
    """절전 대기 시간을 처음부터 다시 셈(5분마다 불리므로 절전 시간이 5분보다 길면 PC가 잠들지 않음).
    원격 세션이 주말·밤에도 PC에 연결할 수 있게(10-23 사용자 요청). 화면 꺼짐은 막지 않음. 설정은 바꾸지 않음."""
    if os.name != "nt":
        return
    try:
        import ctypes
        ctypes.windll.kernel32.SetThreadExecutionState(0x00000001)   # ES_SYSTEM_REQUIRED (한 번, 대기 타이머 초기화)
    except Exception:
        pass


def spawn(p):
    """요청 하나를 따로 띄운 프로세스에서 실행하고 감시 작업은 바로 끝냄(10-27: 감시 작업이 긴 전수시험을 붙잡고 있어
    작업 스케줄러가 새로 띄우지 않아 다른 요청이 몇 시간 줄 서던 문제). 처리 중 표시는 <요청>.json → .run 이름 바꾸기."""
    run_p = p[:-5] + ".run"
    try:
        os.replace(p, run_p)
    except OSError:
        return
    args = [sys.executable, os.path.abspath(__file__), "--do", run_p]
    base = (NO_WINDOW | 0x00000008 | 0x00000200) if os.name == "nt" else 0      # 창 없음·분리·새 그룹
    for fl in ((base | 0x01000000, base) if os.name == "nt" else (0,)):          # 작업 스케줄러 작업 개체에서 빠져나옴(안 되면 그냥)
        try:
            subprocess.Popen(args, cwd=HERE, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL, creationflags=fl, close_fds=True)
            return
        except OSError:
            continue
    os.replace(run_p, p)                      # 못 띄웠으면 요청을 되돌려 다음에


def main():
    keep_awake()
    if not os.path.isdir(REQ):
        return
    os.makedirs(OUT, exist_ok=True)
    if "--do" in sys.argv:
        handle(sys.argv[sys.argv.index("--do") + 1])
        return
    lock = os.path.join(ROOT, "수집함", "_실행중.lock")
    busy = os.path.exists(lock) and dt.datetime.now().timestamp() - os.path.getmtime(lock) < 4 * 3600
    for p in sorted(glob.glob(os.path.join(REQ, "*.json"))):
        try:
            job = json.load(open(p, encoding="utf-8")).get("작업")
        except Exception:
            job = None
        if busy and job in ("전수시험", "시험", "자동실행"):
            continue                          # 정기 실행 중이면 남겨 두고 다음 5분에
        spawn(p)


def handle(p):
    if True:
        rid = os.path.splitext(os.path.basename(p))[0]
        try:
            req = json.load(open(p, encoding="utf-8"))
            job = req.get("작업")
        except Exception:
            req, job = {}, None
        try:
            os.remove(p)                      # 같은 요청을 두 번 하지 않게 먼저 지움
        except OSError:
            return
        with open(os.path.join(OUT, rid + ".txt"), "w", encoding="utf-8") as out:
            out.write(f"요청 {rid} 작업={job} 시작 {dt.datetime.now():%Y-%m-%d %H:%M:%S}\n")
            try:
                if job == "자동실행":
                    run([sys.executable, os.path.join(HERE, "auto_run.py")], out, HERE, 6 * 3600)
                elif job == "시험":
                    selftest(out)
                elif job == "중복정리":                     # 수집함 'name(1)' 똑같은 사본 → _중복(지우지 않음), 대기열에서 뺌
                    run([sys.executable, os.path.join(HERE, "dedupe_inbox.py"), "--dest", os.path.join(ROOT, "수집함"), "--apply"], out, HERE, 3600)
                elif job == "전수시험":                     # 감사·검토 대상 전부 정답 DSD와 대조(full_test.py, 결과 _요청결과\전수시험_*.md)
                    only = req.get("대상")
                    run([sys.executable, os.path.join(HERE, "full_test.py")] + (["--only", only] if only else []), out, HERE, 6 * 3600)
                elif job == "재분류":                       # _미분류를 지금 규칙으로 다시 분류(옮기기만, 지우지 않음). {"확인만": true}면 몇 개 옮길지만
                    run([sys.executable, os.path.join(HERE, "ingest_spc.py"), "--config", os.path.join(HERE, "spc_mail_config.json"), "--reroute"]
                        + (["--dry-run"] if req.get("확인만") else []), out, HERE, 3600)
                elif job == "리뷰수집":                     # 리뷰어(설정 reviewers) 메일 → 수집함\_리뷰메일\<SPC>\ (PC 밖으로 안 보냄)
                    run([sys.executable, os.path.join(HERE, "review_mail.py")], out, HERE, 3 * 3600)
                else:
                    out.write("허용되지 않은 작업 - 무시\n")
            except Exception:
                out.write(traceback.format_exc())
            out.write(f"\n끝 {dt.datetime.now():%Y-%m-%d %H:%M:%S}\n")


if __name__ == "__main__":
    main()
