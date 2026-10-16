"""Master store + inbox <-> Cloudflare R2, for .github/workflows/build.yml (via the admin Functions, Bearer BUILD_TOKEN).

    ci_store.py pull  <home> <inbox>   store/current -> <home>; inbox files -> <inbox>/<slot>.xlsx (+ _options.json)
    ci_store.py push  <home>           upload the updated store files to staging (nothing becomes current yet)
    ci_store.py promote <home>         staging -> current (previous current kept as a version), processed inbox removed
Needs ADMIN_URL, BUILD_TOKEN. Data only ever lands in the folders given (the runner's temp dir)."""
import json, os, sys, urllib.request, urllib.error

URL = os.environ["ADMIN_URL"].rstrip("/") + "/api/build/"
H = {"authorization": "Bearer " + os.environ["BUILD_TOKEN"], "user-agent": "logistik-sp-build/1.0 (GitHub Actions)"}
PUSH = ["data/rilis.csv.gz", "data/so.csv.gz", "data/target_ekspeditur.csv.gz", "data/target_distributor.csv.gz", "config/ekspeditur_map.json",
        "data/gudang.csv.gz", "data/kapal.csv.gz"]   # optional files are skipped when absent

def call(path, method="GET", data=None, ctype=None, timeout=600):
    h = dict(H)
    if ctype: h["content-type"] = ctype
    return urllib.request.urlopen(urllib.request.Request(URL + path, data=data, method=method, headers=h), timeout=timeout)

def download(path, dest):
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    with call(path) as r, open(dest, "wb") as f:
        while True:
            b = r.read(1 << 20)
            if not b: break
            f.write(b)
    return os.path.getsize(dest)

def main():
    cmd, home = sys.argv[1], sys.argv[2]
    if cmd == "pull":
        inbox = sys.argv[3]; os.makedirs(inbox, exist_ok=True)
        files = json.load(call("store"))["files"]
        if not any(f["path"] == "data/rilis.csv.gz" for f in files): sys.exit("GAGAL: master awal belum dimuat di R2 (admin, bagian 0).")
        for f in files: print("store", f["path"], download("store/" + f["path"], os.path.join(home, f["path"])), "byte")
        ib = json.load(call("inbox"))
        for f in ib["files"]: print("inbox", f["slot"], download("inbox/" + f["slot"], os.path.join(inbox, f["slot"] + ".xlsx")), "byte")
        json.dump(ib.get("options") or {}, open(os.path.join(inbox, "_options.json"), "w"))
        json.dump([{"slot": f["slot"], "etag": f["etag"]} for f in ib["files"]], open(os.path.join(home, "_inbox_pulled.json"), "w"))
        print("slot di inbox:", ", ".join(f["slot"] for f in ib["files"]) or "(kosong: build ulang dari master tersimpan)")
    elif cmd == "push":
        for p in PUSH:
            if p in ("data/gudang.csv.gz", "data/kapal.csv.gz") and not os.path.exists(os.path.join(home, p)): continue
            b = open(os.path.join(home, p), "rb").read()
            call("staging/" + p, "PUT", b, "application/octet-stream"); print("staging", p, len(b), "byte")
    elif cmd == "promote":
        meta = json.load(open(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "meta.json")))
        body = {"asOf": meta.get("asOf"), "build": meta.get("build"), "inbox": json.load(open(os.path.join(home, "_inbox_pulled.json")))}
        print(json.load(call("promote", "POST", json.dumps(body).encode(), "application/json")))
    else: sys.exit("perintah tidak dikenal")
try:
    main()
except urllib.error.HTTPError as e:
    print("GAGAL: server menjawab HTTP %s untuk %s" % (e.code, e.url.replace(URL, "/api/build/")), file=sys.stderr); sys.exit(1)
