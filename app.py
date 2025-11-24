# app.py - Gmail Scanner Web UI (ready for Render/gunicorn)
# Exposes `app` for gunicorn: run with `gunicorn app:app --bind 0.0.0.0:$PORT --worker-class gthread --threads 4 --timeout 600`

from flask import Flask, render_template, request, jsonify, Response, send_file, abort
import os, subprocess, time
from threading import Lock
from collections import deque

# minimal: no heavy work on import time
app = Flask(__name__)

# Config (override via env)
GRAB_SCRIPT = os.environ.get("GRAB_SCRIPT", "grab.py")
GRAB_SCRIPT = GRAB_SCRIPT if os.path.isabs(GRAB_SCRIPT) else os.path.join(os.getcwd(), GRAB_SCRIPT)
OUTPUT_XLSX = os.environ.get("OUTPUT_XLSX", "hasil_subject.xlsx")
LOGFILE = os.environ.get("LOGFILE", "/tmp/grab_run.log")
PY_EXEC = os.environ.get("PYTHON_EXEC", "python3")
LOG_TOKEN = os.environ.get("LOG_TOKEN", "")  # optional simple protection for /log endpoints

# subprocess holder + lock
proc = None
proc_lock = Lock()

# ---------- helpers ----------
def is_running():
    with proc_lock:
        return proc is not None and proc.poll() is None

def start_process():
    global proc
    with proc_lock:
        if proc is not None and proc.poll() is None:
            return False, f"Already running pid={proc.pid}"
        # ensure logfile exists and open in append binary, unbuffered
        os.makedirs(os.path.dirname(LOGFILE) or ".", exist_ok=True)
        lf = open(LOGFILE, "ab", buffering=0)
        # launch unbuffered python
        cmd = [PY_EXEC, "-u", GRAB_SCRIPT]
        proc = subprocess.Popen(cmd, stdout=lf, stderr=subprocess.STDOUT)
        return True, f"Started pid={proc.pid}"

def stop_process():
    global proc
    with proc_lock:
        if proc is None or proc.poll() is not None:
            return False, "Not running"
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()
        return True, "Stopped"

def tail_lines(path, n=200):
    if not os.path.exists(path):
        return ""
    dq = deque(maxlen=n)
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            for line in f:
                dq.append(line.rstrip("\n"))
    except Exception as e:
        return f"Error reading log: {e}"
    return "\n".join(dq)

# ---------- routes ----------
@app.route("/")
def index():
    # simple UI loader; expects templates/index.html present in repo
    return render_template("index.html")

@app.route("/start", methods=["POST"])
def start():
    ok, msg = start_process()
    return jsonify({"ok": ok, "msg": msg})

@app.route("/stop", methods=["POST"])
def stop():
    ok, msg = stop_process()
    return jsonify({"ok": ok, "msg": msg})

@app.route("/status")
def status():
    running = False; pid = None; exitcode = None
    with proc_lock:
        if proc is not None:
            if proc.poll() is None:
                running = True; pid = proc.pid
            else:
                exitcode = proc.returncode
    exists = os.path.exists(LOGFILE)
    size = os.path.getsize(LOGFILE) if exists else 0
    return jsonify({
        "running": running,
        "pid": pid,
        "exitcode": exitcode,
        "log_exists": exists,
        "log_size": size,
        "grab_script": GRAB_SCRIPT,
        "output_xlsx": OUTPUT_XLSX
    })

@app.route("/log")
def log_tail():
    token = request.args.get("token", "")
    if LOG_TOKEN and token != LOG_TOKEN:
        abort(401)
    try:
        lines = int(request.args.get("lines", "200"))
    except:
        lines = 200
    return Response(tail_lines(LOGFILE, lines), mimetype="text/plain; charset=utf-8")

@app.route("/download")
def download():
    if not os.path.exists(OUTPUT_XLSX):
        return ("Not found", 404)
    return send_file(OUTPUT_XLSX, as_attachment=True)

@app.route("/stream")
def stream():
    # simple SSE generator that yields new lines (non-blocking for threaded/gevent workers)
    # starts by seeking to end so client gets new events only.
    def generate():
        # create logfile if missing
        if not os.path.exists(LOGFILE):
            open(LOGFILE, "w", encoding="utf-8").close()
        with open(LOGFILE, "r", encoding="utf-8", errors="replace") as f:
            f.seek(0, os.SEEK_END)
            while True:
                line = f.readline()
                if line:
                    yield f"data: {line.rstrip()}\\n\\n"
                else:
                    # heartbeat; check process end
                    time.sleep(0.5)
                    with proc_lock:
                        p = proc
                    if p is not None and p.poll() is not None:
                        yield f"event: done\\ndata: Process finished with exit code {p.returncode}\\n\\n"
                        break
    return Response(generate(), mimetype="text/event-stream")

# ---------- small health endpoint ----------
@app.route("/health")
def health():
    return jsonify({"status":"ok"})

# ---------- run (debug only) ----------
if __name__ == "__main__":
    # safe default for local dev
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), debug=True)
