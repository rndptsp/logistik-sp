"""
Build assets/kec/<province>.js: kecamatan borders per kabupaten/kota, for the district map on the district page
(loaded only when a district page is opened). Run in GitHub Actions (workflow "Header photos & map data", mode kecamatan),
because the sources are not reachable from every network. Needs no extra packages.

  python3 tools/make_kecamatan.py

Sources (open data, CC BY 4.0): geoBoundaries IDN ADM3 (kecamatan) and ADM2 (kabupaten/kota, the same file as
assets/districts.js, so district names match). Each kecamatan goes to the district that contains its centroid.
Coordinates use the same projection as districts.js (200 px/degree).
"""
import json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_districts import ADM2, PROV, PROV_NAME, SUMATRA, ROOT, fetch, polys, area, centroid, inside, simplify, proj

API = "https://www.geoboundaries.org/api/current/gbOpen/IDN/ADM3/"
TOL = 0.0025       # degrees (~280 m): finer than the province map, the district map is zoomed in
MIN_AREA = 0.00015


def shapes(features, tol, min_area):
    """[(name, centroid, svg path, w, h)] for features inside Sumatra."""
    out = []
    for f in features:
        rings = [p[0] for p in polys(f["geometry"])]
        big = max(rings, key=lambda r: abs(area(r)))
        c = centroid(big)
        if not (94 < c[0] < 109 and -7 < c[1] < 7):
            continue
        d, allxy = [], []
        for r in rings:
            if r is not big and abs(area(r)) < min_area:
                continue
            s = simplify(r[:-1] if r[0] == r[-1] else r, tol)
            if len(s) < 3:
                continue
            xy = [proj(p) for p in s]
            allxy += xy
            d.append("M" + "L".join(f"{x},{y}" for x, y in xy) + "Z")
        if allxy:
            out.append((f["properties"]["shapeName"], c, rings, "".join(d), allxy))
    return out


ADM3_URLS = [   # tried in order; the API answer names the current file
    lambda: fetch(API)["simplifiedGeometryGeoJSON"],
    lambda: "https://media.githubusercontent.com/media/wmgeolab/geoBoundaries/main/releaseData/gbOpen/IDN/ADM3/geoBoundaries-IDN-ADM3_simplified.geojson",
    lambda: "https://raw.githubusercontent.com/wmgeolab/geoBoundaries/main/releaseData/gbOpen/IDN/ADM3/geoBoundaries-IDN-ADM3_simplified.geojson",
]


def load_adm3():
    errs = []
    for get_url in ADM3_URLS:
        try:
            data = fetch(get_url())
            if data.get("features"):
                return data["features"]
            errs.append("no features")
        except Exception as e:   # noqa: BLE001 - report every source, then give up
            errs.append(f"{type(e).__name__}: {e}"[:200])
    print("::error::ADM3 tidak bisa diunduh: " + " | ".join(errs))
    sys.exit(1)


def main():
    adm3 = load_adm3()
    adm2 = [f for f in fetch(ADM2)["features"] if not f["properties"]["shapeName"].startswith("Danau")]
    provs = []
    for f in fetch(PROV)["features"]:
        name = PROV_NAME.get(f["properties"]["Propinsi"], f["properties"]["Propinsi"])
        if name in SUMATRA:
            provs.append((name, [p[0] for p in polys(f["geometry"])]))
    # district (ADM2) rings and province, keyed by the same name as in districts.js
    dists = []
    for f in adm2:
        rings = [p[0] for p in polys(f["geometry"])]
        c = centroid(max(rings, key=lambda r: abs(area(r))))
        if not (94 < c[0] < 109 and -7 < c[1] < 7):
            continue
        prov = next((n for n, rs in provs if any(inside(c, r) for r in rs)), None)
        dists.append((f["properties"]["shapeName"], prov, rings, c))

    def district_of(pt):
        for n, prov, rings, _ in dists:
            if prov and any(inside(pt, r) for r in rings):
                return n, prov
        return None

    out, lost = {}, 0
    for name, c, rings, d, allxy in shapes(adm3, TOL, MIN_AREA):
        hit = district_of(c)
        if not hit:   # centroid on a border/sea: try points pulled toward the centroid
            big = max(rings, key=lambda r: abs(area(r)))
            for p in big[::max(1, len(big) // 30)]:
                hit = district_of((p[0] * 0.6 + c[0] * 0.4, p[1] * 0.6 + c[1] * 0.4))
                if hit:
                    break
        if not hit:
            lost += 1
            continue
        dn, prov = hit
        cx, cy = proj(c)
        xs, ys = [x for x, _ in allxy], [y for _, y in allxy]
        out.setdefault(prov, {}).setdefault(dn, []).append({"n": name, "d": d, "cx": cx, "cy": cy,
                                                            "w": round(max(xs) - min(xs), 1), "h": round(max(ys) - min(ys), 1)})
    os.makedirs(os.path.join(ROOT, "assets", "kec"), exist_ok=True)
    for prov, ds in sorted(out.items()):
        res = {}
        for dn, ks in ds.items():
            nums = [float(v) for k in ks for v in re.split(r"[MLZ,]", k["d"]) if v]
            xs, ys = nums[0::2], nums[1::2]
            pad = max(max(xs) - min(xs), max(ys) - min(ys)) * 0.03
            res[dn] = {"vb": [round(min(xs) - pad, 1), round(min(ys) - pad, 1), round(max(xs) - min(xs) + 2 * pad, 1), round(max(ys) - min(ys) + 2 * pad, 1)],
                       "d": sorted(ks, key=lambda k: k["n"])}
        fn = re.sub(r"[^a-z0-9]+", "-", prov.lower()) + ".js"
        path = os.path.join(ROOT, "assets", "kec", fn)
        with open(path, "w", encoding="utf-8") as fh:
            fh.write("/* Kecamatan borders per kabupaten/kota in " + prov + ", from geoBoundaries IDN ADM3 (CC BY 4.0), simplified. "
                     "Built by tools/make_kecamatan.py; same projection as districts.js. */\n")
            fh.write("(window.KEC = window.KEC || {})[" + json.dumps(prov) + "] = " + json.dumps(res, separators=(",", ":"), ensure_ascii=False) + ";\n")
        print(f"  {prov}: {len(res)} kab/kota, {sum(len(v) for v in res.values())} kecamatan, {os.path.getsize(path) / 1e3:.0f} KB")
    print(f"tanpa kab/kota: {lost}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:   # surface the reason as a run annotation (job logs are not always readable)
        print(f"::error::{type(e).__name__}: {str(e)[:300]}")
        raise
