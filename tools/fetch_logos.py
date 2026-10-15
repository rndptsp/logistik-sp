"""Regional emblems (lambang provinsi / kabupaten / kota) from Wikimedia Commons -> assets/logo/*.png + assets/logo/logo.js.
Runs in GitHub Actions (workflow "Header photos & map data", mode logos).

    python3 tools/fetch_logos.py

Names come from assets/districts.js (same names as the district maps). For each region the Commons search is
"Lambang <Provinsi|Kabupaten|Kota> <name>"; the best-scoring file whose title names the region and an emblem is taken,
as a 256 px PNG. Indonesian government emblems are public domain on Commons (PD-IDGov); the licence and source page are
kept for the credits page (#/kredit). Regions with no confident match are left out (the header then shows no logo).
"""
import io, json, os, re, time, urllib.parse, urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "assets", "logo")
UA = {"User-Agent": "logistik-sp-dashboard/1.0 (internal company dashboard)"}
BAD = re.compile(r"kecamatan|desa|nagari|dprd|polres|polda|kodim|partai|universitas|sekolah|rsud|pkk|korpri|bpbd|old|lama|1950|flag|bendera|map|peta|locator|location", re.I)
EMB = re.compile(r"lambang|logo|coat.of.arms|seal|emblem", re.I)


def api(p):
    for i in range(6):
        try:
            time.sleep(1.5)
            url = "https://commons.wikimedia.org/w/api.php?" + urllib.parse.urlencode(p)
            return json.loads(urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=60).read())
        except urllib.error.HTTPError as e:
            if e.code != 429:
                raise
            time.sleep(30 * (i + 1))
    return {}


def get(url):
    for i in range(5):
        try:
            time.sleep(1)
            return urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=60).read()
        except urllib.error.HTTPError as e:
            if e.code != 429:
                raise
            time.sleep(30 * (i + 1))
    raise RuntimeError("429")


def words(s):
    return [w for w in re.sub(r"[^a-z0-9 ]", " ", s.lower()).split() if w not in ("kota", "kabupaten", "provinsi", "kab")]


def best(kind, name):
    q = f"Lambang {kind} {name}"
    r = api({"action": "query", "format": "json", "generator": "search", "gsrsearch": q, "gsrnamespace": 6, "gsrlimit": 20,
             "prop": "imageinfo", "iiprop": "url|extmetadata|mime", "iiurlwidth": 256})
    need = words(name)
    cands = []
    for p in r.get("query", {}).get("pages", {}).values():
        t = p["title"][5:]
        tl = re.sub(r"[_\-]", " ", t.lower())
        ii = (p.get("imageinfo") or [{}])[0]
        if not EMB.search(tl) or BAD.search(tl) or not all(w in words(tl) for w in need) or "thumburl" not in ii:
            continue
        extra = len(words(re.sub(r"\.(svg|png|jpe?g|gif)$", "", tl))) - len(need) - 1   # words beyond "lambang <name>"
        score = (3 if "lambang" in tl else 1) + (2 if kind.lower() in tl else 0) + (1 if tl.endswith(".svg") else 0) - max(0, extra) - p.get("index", 0) * 0.05
        if kind == "Kabupaten" and re.search(r"\bkota\b", tl) or kind == "Kota" and "kabupaten" in tl:
            continue
        cands.append((score, t, ii))
    return max(cands, key=lambda c: c[0]) if cands else None


def main():
    s = open(os.path.join(ROOT, "assets", "districts.js"), encoding="utf-8").read()
    D = json.loads(s[s.index("=") + 1:].strip().rstrip(";"))
    slug = lambda x: re.sub(r"[^a-z0-9]+", "-", x.lower()).strip("-")
    prov_name = {"RIAU DARATAN": "Riau", "KEPULAUAN RIAU": "Kepulauan Riau", "BANGKA BELITUNG": "Kepulauan Bangka Belitung"}
    jobs = []
    for prov, G in D.items():
        jobs.append(("prov-" + slug(prov), "Provinsi", prov_name.get(prov, prov.title()), prov, None))
        for g in G["d"]:
            kota = g["n"].startswith("Kota ")
            jobs.append(("dist-" + slug(prov) + "-" + slug(g["n"]), "Kota" if kota else "Kabupaten", g["n"][5:] if kota else g["n"], prov, g["n"]))
    os.makedirs(OUT, exist_ok=True)
    meta, miss = {}, []
    for key, kind, name, prov, shape in jobs:
        b = best(kind, name)
        if not b:
            miss.append(f"{kind} {name}"); continue
        _, title, ii = b
        md = ii.get("extmetadata", {})
        clean = lambda k: re.sub(r"\s+", " ", re.sub("<[^>]+>", "", md.get(k, {}).get("value", ""))).strip()
        try:
            png = get(ii["thumburl"])
        except Exception:   # noqa: BLE001
            miss.append(f"{kind} {name} (unduh gagal)"); continue
        fn = key + ".png"
        open(os.path.join(OUT, fn), "wb").write(png)
        meta[key] = {"f": fn, "t": f"Lambang {kind} {name}", "a": (clean("Artist") or "Pemerintah daerah")[:80],
                     "l": clean("LicenseShortName") or "Public domain", "u": ii.get("descriptionurl", ""), "p": prov, "n": shape}
        print("  ok", key, "<-", title[:70])
    with open(os.path.join(OUT, "logo.js"), "w", encoding="utf-8") as f:
        f.write("/* Regional emblems for page headers (see applyPageBg in app.js). Generated by tools/fetch_logos.py; credits at #/kredit. */\n")
        f.write("window.LOGO = " + json.dumps(meta, ensure_ascii=False, indent=0) + ";\n")
    print(f"{len(meta)} logo, {len(miss)} tidak ketemu: " + "; ".join(miss))


if __name__ == "__main__":
    main()
