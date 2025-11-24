# app.py
import os
import time
import subprocess
from threading import Lock
from flask import Flask, render_template, request, jsonify, send_file, Response

# PATH TO YOUR EXISTING SCRIPT (from your session)
GRAB_SCRIPT = "/mnt/data/grab.py"           # <- existing script path (do not change unless you moved it)
OUTPUT_XLSX = "/mnt/data/hasil_subject.xlsx"  # file produced by the script
LOGFILE = "/tmp/grab_run.log"               # live log file used by web UI

app = Flask(__name__)
proc = None           # will hold subprocess.Popen instance
proc_lock = Lock()


def start_scan_process():
    """
    Start the external grab.py as a background process that writes stdout/stderr to LOGFILE.
    If already running, returns False.
    """
    global proc
    with proc_lock:
        if proc is not None and proc.poll() is None:
            # already running
            return False, "Process already running"
        # ensure logfile dir exists
        logfile_dir = os.path.dirname(LOGFILE)
        os.makedirs(logfile_dir or ".", exist_ok=True)
        # open logfile
        lf = open(LOGFILE, "wb")
        # start process: use the same python executable
        cmd = [os.environ.get("PYTHON_EXEC", "python3"), GRAB_SCRIPT]
        # spawn process, redirect stdout+stderr to logfile
        proc = subprocess.Popen(cmd, stdout=lf, stderr=subprocess.STDOUT)
        return True, f"Started process pid={proc.pid}"


def stop_scan_process():
    global proc
    with proc_lock:
        if proc is None or proc.poll() is not None:
            return False, "No running process"
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()
        return True, "Stopped"


def stream_log():
    """
    Server-Sent Events generator that tails the logfile.
    It yields lines as 'data: <line>\\n\\n'
    """
    # if no logfile yet, create empty one
    if not os.path.exists(LOGFILE):
        open(LOGFILE, "w", encoding="utf-8").close()

    def generate():
        with open(LOGFILE, "r", encoding="utf-8", errors="replace") as f:
            # seek to end to stream only new lines — but we'll start from beginning so user sees history:
            f.seek(0, os.SEEK_SET)
            while True:
                line = f.readline()
                if line:
                    yield f"data: {line.rstrip()}\\n\\n"
                else:
                    # also send simple status events every second (heartbeat)
                    # but check process status and send special event when finished
                    time.sleep(0.8)
                    with proc_lock:
                        p = proc
                    if p is not None and p.poll() is not None:
                        # process finished; send final event and stop generator
                        yield f"event: done\\ndata: Process finished with exit code {p.returncode}\\n\\n"
                        break
    return Response(generate(), mimetype="text/event-stream")


@app.route("/")
def index():
    # report running state
    running = False
    pid = None
    exitcode = None
    with proc_lock:
        if proc is not None:
            if proc.poll() is None:
                running = True
                pid = proc.pid
            else:
                exitcode = proc.returncode
    # exists output?
    output_exists = os.path.exists(OUTPUT_XLSX)
    output_size = os.path.getsize(OUTPUT_XLSX) if output_exists else 0
    return render_template("index.html",
                           running=running, pid=pid, exitcode=exitcode,
                           output_exists=output_exists, output_size=output_size)


@app.route("/start", methods=["POST"])
def start():
    ok, msg = start_scan_process()
    return jsonify({"ok": ok, "msg": msg})


@app.route("/stop", methods=["POST"])
def stop():
    ok, msg = stop_scan_process()
    return jsonify({"ok": ok, "msg": msg})


@app.route("/stream")
def stream():
    return stream_log()


@app.route("/download")
def download():
    if not os.path.exists(OUTPUT_XLSX):
        return ("Not found", 404)
    # send_file will stream the file
    return send_file(OUTPUT_XLSX, as_attachment=True)

if __name__ == "__main__":
    # debug mode for development — in production use WSGI server (gunicorn/uwsgi)
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), debug=True)
