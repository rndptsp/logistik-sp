"""Search Wikimedia Commons for header-photo candidates (runs in GitHub Actions, see .github/workflows/bg.yml).

    python3 tools/bg_search.py tools/bg_queries.json > tools/bg_candidates.json

Lists only landscape photos >= 1600 px with a free licence (CC0 / PD / CC BY / CC BY-SA), with a 330 px preview URL,
so a person can pick one per page key and put it in tools/bg_photos.json for tools/fetch_bg.py.
"""
import json, re, sys, time, urllib.parse, urllib.request

UA = {"User-Agent": "logistik-sp-dashboard/1.0 (internal company dashboard)"}
OK = re.compile(r"^(CC0|Public domain|PD|CC BY(-SA)? ?\d)", re.I)


def api(p):
    for i in range(6):
        try:
            time.sleep(2)
            url = "https://commons.wikimedia.org/w/api.php?" + urllib.parse.urlencode(p)
            return json.loads(urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=60).read())
        except urllib.error.HTTPError as e:
            if e.code != 429:
                raise
            time.sleep(30 * (i + 1))
    return {}


def main():
    out = {}
    for key, queries in json.load(open(sys.argv[1], encoding="utf-8")).items():
        seen, res = set(), []
        for q in queries:
            r = api({"action": "query", "format": "json", "generator": "search", "gsrsearch": q + " filetype:bitmap",
                     "gsrnamespace": 6, "gsrlimit": 25, "prop": "imageinfo", "iiprop": "url|size|extmetadata", "iiurlwidth": 330})
            for p in sorted(r.get("query", {}).get("pages", {}).values(), key=lambda p: p.get("index", 0)):
                ii = (p.get("imageinfo") or [{}])[0]
                md = ii.get("extmetadata", {})
                clean = lambda k: re.sub(r"\s+", " ", re.sub("<[^>]+>", "", md.get(k, {}).get("value", ""))).strip()
                lic = clean("LicenseShortName")
                if p["title"] in seen or ii.get("width", 0) < 1600 or ii["width"] < ii["height"] * 1.2 or not OK.match(lic):
                    continue
                seen.add(p["title"])
                res.append({"file": p["title"], "w": ii["width"], "h": ii["height"], "lic": lic, "a": clean("Artist")[:60],
                            "d": clean("ImageDescription")[:120], "prev": ii.get("thumburl", ""), "page": ii.get("descriptionurl", "")})
        out[key] = res[:12]
        print(key, len(res), file=sys.stderr)
    json.dump(out, sys.stdout, indent=1, ensure_ascii=False)


if __name__ == "__main__":
    main()
