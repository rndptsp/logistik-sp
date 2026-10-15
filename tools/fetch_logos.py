"""Regional emblems (lambang provinsi / kabupaten / kota) from Wikimedia Commons -> assets/logo/*.png + assets/logo/logo.js.
Runs in GitHub Actions (workflow "Header photos & map data", mode logos).

    python3 tools/fetch_logos.py

Names come from assets/districts.js (same names as the district maps). The emblem is taken from the files used on the
region's Indonesian Wikipedia article ("Kabupaten X", "Kota Y", province name), read 40 articles per request so the run
stays within Wikimedia rate limits; a file named "Lambang/Logo/Coat of arms ..." is taken as a 256 px PNG. Indonesian government emblems are public domain on Commons (PD-IDGov); the licence and source page are
kept for the credits page (#/kredit). Regions with no confident match are left out (the header then shows no logo).
"""
import io, json, os, re, time, urllib.parse, urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "assets", "logo")
UA = {"User-Agent": "logistik-sp-dashboard/1.0 (internal company dashboard)"}
BAD = re.compile(r"kecamatan|desa|nagari|dprd|polres|polda|kodim|partai|universitas|sekolah|rsud|pkk|korpri|bpbd|\bold\b|\blama\b|flag|bendera|\bmap\b|\bpeta\b|locator|location|\b1[89]\d\d\b|kolonial|hindia|oosthaven", re.I)
EMB = re.compile(r"lambang|logo|coat.of.arms|seal|emblem", re.I)


WIKI = "https://id.wikipedia.org/w/api.php?"


def api(p, base=WIKI):
    for i in range(6):
        try:
            time.sleep(2)
            return json.loads(urllib.request.urlopen(urllib.request.Request(base + urllib.parse.urlencode(p), headers=UA), timeout=60).read())
        except urllib.error.HTTPError as e:
            if e.code != 429:
                raise
            time.sleep(20 * (i + 1))
    return {}


def get(url):
    for i in range(4):
        try:
            time.sleep(0.5)
            return urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=60).read()
        except urllib.error.HTTPError as e:
            if e.code != 429:
                raise
            time.sleep(15 * (i + 1))
    raise RuntimeError("429")


EN = {"Sumatera Barat": ["West Sumatra"], "Sumatera Utara": ["North Sumatra"], "Sumatera Selatan": ["South Sumatra"],
      "Kepulauan Riau": ["Riau Islands"], "Kepulauan Bangka Belitung": ["Bangka Belitung Islands", "Bangka Belitung"],
      "Labuhan Batu": ["Labuhanbatu"], "Banyu Asin": ["Banyuasin"], "Sawah Lunto": ["Sawahlunto"], "Tulangbawang": ["Tulang Bawang"]}


def words(s):
    return [w for w in re.sub(r"[^a-z0-9 ]", " ", s.lower()).split() if w not in ("kota", "kabupaten", "provinsi", "kab", "file", "berkas")]


def pick(files, kind, name):
    """Best emblem among the files used on the region's Wikipedia article."""
    best = None
    for f in files:
        tl = re.sub(r"[_\-]", " ", f.lower())
        if not EMB.search(tl) or BAD.search(tl):
            continue
        joined = "".join(words(tl))
        if not any("".join(words(n)) in joined for n in [name] + EN.get(name, [])):   # the file must name this region (spacing-insensitive)
            continue
        if kind != "Provinsi" and re.search(r"islands|province|provinsi|kepulauan", tl) and "kepulauan" not in name.lower():
            continue   # e.g. "Bangka" must not take the Bangka Belitung province emblem
        if kind == "Provinsi" and re.search(r"kabupaten|\bkota\b", tl) or kind == "Kabupaten" and re.search(r"\bkota\b", tl) and "kabupaten" not in tl:
            continue
        score = (3 if "lambang" in tl else 2) + (1 if tl.endswith(".svg") else 0)
        if best is None or score > best[0]:
            best = (score, f)
    return best[1] if best else None


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
    title = lambda kind, name: name if kind == "Provinsi" else f"{kind} {name}"
    files = {}
    for i in range(0, len(jobs), 40):   # article -> files used on it, 40 articles per request
        chunk = jobs[i:i + 40]
        cont = {}
        while True:
            r = api(dict({"action": "query", "format": "json", "titles": "|".join(title(j[1], j[2]) for j in chunk),
                          "prop": "images", "imlimit": "max", "redirects": 1}, **cont))
            q = r.get("query", {})
            alias = {x["to"]: x["from"] for x in q.get("redirects", []) + q.get("normalized", [])}
            for pg in q.get("pages", {}).values():
                t = pg.get("title", "")
                while t in alias:
                    t = alias[t]
                files.setdefault(t, []).extend(im["title"] for im in pg.get("images", []))
            if "continue" not in r:
                break
            cont = r["continue"]
    chosen = {}
    for key, kind, name, prov, shape in jobs:
        f = pick(files.get(title(kind, name), []), kind, name)
        if f:
            chosen[key] = (f, kind, name, prov, shape)
        else:
            miss.append(f"{kind} {name}")
    info = {}
    fl = sorted({v[0] for v in chosen.values()})
    for i in range(0, len(fl), 40):
        r = api({"action": "query", "format": "json", "titles": "|".join(fl[i:i + 40]), "prop": "imageinfo",
                 "iiprop": "url|extmetadata", "iiurlwidth": 256})
        q = r.get("query", {})
        alias = {x["to"]: x["from"] for x in q.get("normalized", [])}
        for pg in q.get("pages", {}).values():
            if pg.get("imageinfo"):
                info[alias.get(pg["title"], pg["title"])] = pg["imageinfo"][0]
    t0 = time.time()
    for key, (f, kind, name, prov, shape) in chosen.items():
        ii = info.get(f)
        if not ii or time.time() - t0 > 900:
            miss.append(f"{kind} {name} (info/waktu)"); continue
        md = ii.get("extmetadata", {})
        clean = lambda k: re.sub(r"\s+", " ", re.sub("<[^>]+>", "", md.get(k, {}).get("value", ""))).strip()
        try:
            png = get(ii.get("thumburl") or ii["url"])
        except Exception:   # noqa: BLE001
            miss.append(f"{kind} {name} (unduh gagal)"); continue
        fn = key + ".png"
        open(os.path.join(OUT, fn), "wb").write(png)
        meta[key] = {"f": fn, "t": f"Lambang {title(kind, name) if kind != 'Provinsi' else 'Provinsi ' + name}", "a": (clean("Artist") or "Pemerintah daerah")[:80],
                     "l": clean("LicenseShortName") or "Public domain", "u": ii.get("descriptionurl", ""), "p": prov, "n": shape}
        print("  ok", key, "<-", f[:70])
    with open(os.path.join(OUT, "logo.js"), "w", encoding="utf-8") as f:
        f.write("/* Regional emblems for page headers (see applyPageBg in app.js). Generated by tools/fetch_logos.py; credits at #/kredit. */\n")
        f.write("window.LOGO = " + json.dumps(meta, ensure_ascii=False, indent=0) + ";\n")
    print(f"{len(meta)} logo, {len(miss)} tidak ketemu: " + "; ".join(miss))


if __name__ == "__main__":
    main()
