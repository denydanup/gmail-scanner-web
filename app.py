# app.py - minimal, guaranteed to expose 'app' for gunicorn
from flask import Flask, render_template, jsonify, request, send_file, Response
import os

app = Flask(__name__)

# simple index (expects templates/index.html)
@app.route("/")
def index():
    return render_template("index.html")

# health check
@app.route("/health")
def health():
    return jsonify({"status": "ok"})

# placeholder start/stop endpoints (implement your subprocess control here)
@app.route("/start", methods=["POST"])
def start():
    # start child process logic...
    return jsonify({"ok": True, "msg": "started"})

@app.route("/stop", methods=["POST"])
def stop():
    # stop child process logic...
    return jsonify({"ok": True, "msg": "stopped"})

# streaming example (SSE) - non-blocking for threaded/gevent workers
@app.route("/stream")
def stream():
    def gen():
        yield "data: welcome\\n\\n"
    return Response(gen(), mimetype="text/event-stream")

# download produced file
@app.route("/download")
def download():
    path = os.environ.get("OUTPUT_XLSX", "/tmp/hasil_subject.xlsx")
    if not os.path.exists(path):
        return "Not found", 404
    return send_file(path, as_attachment=True)

if __name__ == "__main__":
    # debug only; gunicorn will not execute this block
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), debug=True)
