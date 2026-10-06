"""
Build assets/districts.js: kabupaten/kota borders for the 10 Sumatra provinces, for the district maps on
the province and district pages. Run once (the borders rarely change); needs internet, no extra packages.

  python3 tools/make_districts.py

Sources (open data):
  districts: geoBoundaries IDN ADM2 (simplified), github.com/wmgeolab/geoBoundaries, CC BY 4.0
  provinces: github.com/ans-4175/peta-indonesia-geojson (same source as assets/sumatra.js), used to
             assign each district to a province
Output coordinates use the same equirectangular projection as sumatra.js, at 200 px/degree.
"""
import json, os, urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ADM2 = "https://media.githubusercontent.com/media/wmgeolab/geoBoundaries/main/releaseData/gbOpen/IDN/ADM2/geoBoundaries-IDN-ADM2_simplified.geojson"
PROV = "https://raw.githubusercontent.com/ans-4175/peta-indonesia-geojson/master/indonesia-prov.geojson"
PROV_NAME = {"DI. ACEH": "ACEH", "RIAU": "RIAU DARATAN"}   # source name -> dashboard name
SUMATRA = {"ACEH", "SUMATERA UTARA", "SUMATERA BARAT", "RIAU DARATAN", "KEPULAUAN RIAU", "JAMBI",
           "SUMATERA SELATAN", "BENGKULU", "BANGKA BELITUNG", "LAMPUNG"}
S = 200            # px per degree
TOL = 0.006        # simplification tolerance, degrees (~650 m)
MIN_AREA = 0.0006  # drop islands smaller than this (deg², ~7 km²) unless it is the district's largest piece


def fetch(url):
    print("download", url)
    with urllib.request.urlopen(url, timeout=300) as r:
        return json.load(r)


def polys(geom):
    return geom["coordinates"] if geom["type"] == "MultiPolygon" else [geom["coordinates"]]


def area(ring):
    return sum(x0 * y1 - x1 * y0 for (x0, y0), (x1, y1) in zip(ring, ring[1:] + ring[:1])) / 2


def centroid(ring):
    a = area(ring)
    if abs(a) < 1e-12:
        return sum(p[0] for p in ring) / len(ring), sum(p[1] for p in ring) / len(ring)
    cx = cy = 0.0
    for (x0, y0), (x1, y1) in zip(ring, ring[1:] + ring[:1]):
        c = x0 * y1 - x1 * y0
        cx += (x0 + x1) * c; cy += (y0 + y1) * c
    return cx / (6 * a), cy / (6 * a)


def inside(pt, ring):
    x, y, hit = pt[0], pt[1], False
    for (x0, y0), (x1, y1) in zip(ring, ring[1:] + ring[:1]):
        if (y0 > y) != (y1 > y) and x < (x1 - x0) * (y - y0) / (y1 - y0) + x0:
            hit = not hit
    return hit


def simplify(pts, tol):
    """Douglas-Peucker, iterative."""
    if len(pts) < 4:
        return pts
    keep, stack = {0, len(pts) - 1}, [(0, len(pts) - 1)]
    while stack:
        a, b = stack.pop()
        (ax, ay), (bx, by) = pts[a], pts[b]
        dx, dy = bx - ax, by - ay
        n = (dx * dx + dy * dy) ** 0.5 or 1e-12
        best, idx = 0, None
        for i in range(a + 1, b):
            d = abs(dy * (pts[i][0] - ax) - dx * (pts[i][1] - ay)) / n
            if d > best:
                best, idx = d, i
        if idx is not None and best > tol:
            keep.add(idx); stack += [(a, idx), (idx, b)]
    return [pts[i] for i in sorted(keep)]


def proj(p):
    return round((p[0] - 95) * S, 1), round((6.2 - p[1]) * S, 1)


def main():
    provs = []
    for f in fetch(PROV)["features"]:
        name = PROV_NAME.get(f["properties"]["Propinsi"], f["properties"]["Propinsi"])
        if name in SUMATRA:
            provs.append((name, [p[0] for p in polys(f["geometry"])]))

    def province_of(pt):
        for name, rings in provs:
            if any(inside(pt, r) for r in rings):
                return name
        return None

    out = {}
    for f in fetch(ADM2)["features"]:
        if f["properties"]["shapeName"].startswith("Danau"):   # lakes (Danau Toba etc.) are not districts
            continue
        rings = [p[0] for p in polys(f["geometry"])]
        big = max(rings, key=lambda r: abs(area(r)))
        c = centroid(big)
        if not (94 < c[0] < 109 and -7 < c[1] < 7):
            continue
        prov = province_of(c)
        if not prov:   # centroid in the sea or on a border: vote with points pulled toward the centroid
            votes = {}
            for p in big[::max(1, len(big) // 40)]:
                q = province_of((p[0] * 0.7 + c[0] * 0.3, p[1] * 0.7 + c[1] * 0.3))
                if q:
                    votes[q] = votes.get(q, 0) + 1
            prov = max(votes, key=votes.get) if votes else None
        if not prov:
            continue
        d, allxy = [], []
        for r in rings:
            if r is not big and abs(area(r)) < MIN_AREA:
                continue
            s = simplify(r[:-1] if r[0] == r[-1] else r, TOL)
            if len(s) < 3:
                continue
            xy = [proj(p) for p in s]
            allxy += xy
            d.append("M" + "L".join(f"{x},{y}" for x, y in xy) + "Z")
        name = f["properties"]["shapeName"]
        cx, cy = proj(c)
        w = round(max(x for x, _ in allxy) - min(x for x, _ in allxy))
        h = round(max(y for _, y in allxy) - min(y for _, y in allxy))
        out.setdefault(prov, []).append({"n": name, "d": "".join(d), "cx": cx, "cy": cy, "w": w, "h": h})

    res = {}
    for prov, ds in sorted(out.items()):
        nums = [float(v) for x in ds for v in x["d"].replace("M", " ").replace("L", " ").replace("Z", " ").replace(",", " ").split()]
        xs, ys = nums[0::2], nums[1::2]
        pad = 6
        res[prov] = {"vb": [round(min(xs) - pad), round(min(ys) - pad), round(max(xs) - min(xs) + 2 * pad), round(max(ys) - min(ys) + 2 * pad)],
                     "d": sorted(ds, key=lambda x: x["n"])}
        print(f"  {prov}: {len(ds)} kab/kota")
    path = os.path.join(ROOT, "assets", "districts.js")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("/* Kabupaten/kota borders per Sumatra province, from geoBoundaries IDN ADM2 (CC BY 4.0), "
                 "simplified. Built by tools/make_districts.py. Same projection as sumatra.js at 200 px/degree. */\n")
        fh.write("window.DISTRICTS = " + json.dumps(res, separators=(",", ":"), ensure_ascii=False) + ";\n")
    print(f"wrote {path} ({os.path.getsize(path) / 1e3:.0f} KB)")


if __name__ == "__main__":
    main()
