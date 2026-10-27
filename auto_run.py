# -*- coding: utf-8 -*-
"""자동 실행 본체: 메일 수집 → 자동 결산 → (오후) 과거 메일 이어서 가져오기.

SPC수집_자동.bat(작업 스케줄러)이 pythonw로 창 없이 띄운다. 창을 닫아 끊기거나 1시간 제한에 걸리지 않게(10-22).
- 겹침 방지: 수집함\\_실행중.lock 이 4시간 안에 만들어졌으면 건너뜀
- 기록: 수집함\\_자동실행로그.txt (하위 실행 출력도 모두)
"""
import os, sys, shutil, subprocess, datetime as dt, traceback

HERE = os.path.dirname(os.path.abspath(__file__))
BOX = os.path.normpath(os.path.join(HERE, "..", "수집함"))
LOG = os.path.join(BOX, "_자동실행로그.txt")
LOCK = os.path.join(BOX, "_실행중.lock")
AUTO = os.path.normpath(os.path.join(HERE, "..", "표준화", "변환기", "자동결산.py"))
CFG = os.path.join(HERE, "spc_mail_config.json")
PKGS = ["pywin32", "xlrd", "openpyxl", "pyzipper", "pymupdf"]
NO_WINDOW = 0x08000000 if os.name == "nt" else 0


def now():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def python_cmd():
    """창 없는 실행에서도 쓸 python.exe(pythonw 말고) - 하위 실행 출력을 기록에 받기 위해"""
    exe = sys.executable
    if exe.lower().endswith("pythonw.exe"):
        cand = exe[:-len("pythonw.exe")] + "python.exe"
        if os.path.exists(cand):
            exe = cand
    return exe


AGENT_TASK = "SPC_request_watch"


def ensure_agent_task(log):
    """원격 점검 세션의 요청(수집\\_요청)을 5분마다 보는 작업을 한 번 등록(사용자 계정, 관리자 권한 불필요)"""
    if os.name != "nt":
        return
    q = subprocess.run(["schtasks", "/Query", "/TN", AGENT_TASK], capture_output=True, creationflags=NO_WINDOW)
    if q.returncode == 0:
        return
    pyw = sys.executable
    if pyw.lower().endswith("python.exe") and os.path.exists(pyw[:-10] + "pythonw.exe"):
        pyw = pyw[:-10] + "pythonw.exe"
    tr = f'"{pyw}" "{os.path.join(HERE, "pc_agent.py")}"'
    r = subprocess.run(["schtasks", "/Create", "/TN", AGENT_TASK, "/SC", "MINUTE", "/MO", "5", "/TR", tr, "/F"],
                       capture_output=True, creationflags=NO_WINDOW)
    log.write(f"요청 감시 작업 등록: {'성공' if r.returncode == 0 else '실패 ' + r.stderr.decode('cp949', 'replace').strip()}\n")


def main():
    os.makedirs(BOX, exist_ok=True)
    if os.path.exists(LOCK):
        age = dt.datetime.now().timestamp() - os.path.getmtime(LOCK)
        if age < 4 * 3600:
            with open(LOG, "a", encoding="utf-8") as f:
                f.write(f"===== {now()} 이전 실행이 아직 진행 중 - 건너뜀 =====\n")
            return
    with open(LOCK, "w", encoding="utf-8") as f:
        f.write(now())
    env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8")
    try:
        with open(LOG, "a", encoding="utf-8") as log:
            log.write(f"\n===== {now()} (창 없이 실행) =====\n")
            try:
                ensure_agent_task(log)
            except Exception as e:
                log.write(f"요청 감시 작업 등록 오류: {e}\n")
            log.flush()

            def run(args, label=None, timeout=4 * 3600):
                if label:
                    log.write(f"----- {label} -----\n")
                log.flush()
                try:
                    r = subprocess.run(args, cwd=HERE, stdout=log, stderr=subprocess.STDOUT, env=env,
                                       stdin=subprocess.DEVNULL, creationflags=NO_WINDOW, timeout=timeout)
                    return r.returncode
                except subprocess.TimeoutExpired:
                    log.write(f"시간 초과({timeout // 60}분) - 끊고 다음 단계로\n")
                    return -2
                except Exception as e:
                    log.write(f"실행 실패: {type(e).__name__}: {e}\n")
                    return -1
                finally:
                    log.flush()

            uv = shutil.which("uv")
            if uv:
                py = [uv, "run"] + sum((["--with", p] for p in PKGS), []) + ["python"]
            else:
                py = [python_cmd()]
                run(py + ["-m", "pip", "install", "-q", "--disable-pip-version-check"] + PKGS)
            run(py + ["ingest_spc.py", "--days", "14"])
            f200 = os.path.join(HERE, "_과거수집_200.flag")
            if os.path.exists(f200):
                run(py + ["ingest_spc.py", "--days", "200", "--no-queue"], "과거 200일 메일 가져오기")
                os.remove(f200)
            fsc = os.path.join(HERE, "_범위재확인_400.flag")   # 업무배정표 이름 맞추기가 바뀐 뒤(10-24 펀드·리츠) 건너뛴 메일 다시 보기, 한 번
            if os.path.exists(fsc):
                if run(py + ["ingest_spc.py", "--days", "400", "--recheck-scope"], "업무배정표 범위 다시 확인(400일)") == 0:
                    os.remove(fsc)
            run(py + [AUTO, "--config", CFG], "자동 결산", timeout=int(3.5 * 3600))   # 한 SPC가 멈춰도 다음 실행이 막히지 않게(10-24 43건에 2시간 40분 걸림)
            f400 = os.path.join(HERE, "_과거수집_400.flag")
            if os.path.exists(f400) and dt.datetime.now().hour >= 12:
                if run(py + ["ingest_spc.py", "--days", "400", "--no-queue"], "과거 400일 메일 가져오기") == 0:
                    os.remove(f400)
            log.write(f"===== {now()} 끝 =====\n")
    except Exception:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(f"자동 실행 오류 {now()}\n{traceback.format_exc()}\n")
    finally:
        try:
            os.remove(LOCK)
        except OSError:
            pass


if __name__ == "__main__":
    main()
