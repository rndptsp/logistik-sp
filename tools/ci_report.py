"""Tell the admin page how the build went (used by .github/workflows/build.yml).

    python3 tools/ci_report.py running|ok|failed
Needs ADMIN_URL, BUILD_TOKEN, RUN_URL in the environment; sends the build log tail with ok/failed."""
import json, os, sys, urllib.request

state = sys.argv[1]
body = {"state": state, "run_url": os.environ.get("RUN_URL")}
if state != "running":
    log = os.path.join(os.environ.get("RUNNER_TEMP", "."), "build.log")
    body["log"] = open(log, encoding="utf-8", errors="replace").read()[-8000:] if os.path.exists(log) else "(tidak ada log build)"
    if state == "ok":
        meta = json.load(open(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "meta.json")))
        body.update(asOf=meta.get("asOf"), build=meta.get("build"))
req = urllib.request.Request(os.environ["ADMIN_URL"].rstrip("/") + "/api/build/report", data=json.dumps(body).encode(),
                             headers={"authorization": "Bearer " + os.environ["BUILD_TOKEN"], "content-type": "application/json"})
try:
    urllib.request.urlopen(req, timeout=30).read()
    print("status dikirim ke halaman admin:", state)
except Exception as e:   # reporting must never hide the real build result
    print("gagal mengirim status ke halaman admin:", e)
