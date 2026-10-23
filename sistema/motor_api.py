"""
Cliente da API local do motor (gosom/google-maps-scraper) em http://127.0.0.1:8080.

Funcoes copiadas de scripts/scrape.py do Google Maps Scraper Kit
(https://github.com/Mahanaicoach/google-maps-scraper-kit), MIT License,
Copyright (c) 2026 Mahan. Veja CREDITOS.md.
"""
import csv, io, json, math, os, sys, time, urllib.error, urllib.parse, urllib.request

BASE = os.environ.get("SCRAPER_BASE_URL", "http://127.0.0.1:8080")
KEY = os.environ.get("SCRAPER_API_KEY", "")
UA = "clientes-para-sites/1.0 (busca local de empresas sem site)"


def req(method, path, body=None):
    headers = {"Content-Type": "application/json", "User-Agent": UA}
    if KEY:
        headers["X-API-Key"] = KEY
    data = json.dumps(body).encode() if body is not None else None
    r = urllib.request.Request(BASE + path, data=data, headers=headers, method=method)
    with urllib.request.urlopen(r, timeout=60) as resp:
        return resp.status, resp.read()


def geocode(place, with_bbox=False):
    """City/place name -> ('lat','lon') strings via Nominatim, or None.
    with_bbox=True returns ('lat','lon', [south, north, west, east]) instead."""
    q = urllib.parse.urlencode({"format": "json", "limit": 1, "q": place})
    url = f"https://nominatim.openstreetmap.org/search?{q}"
    r = urllib.request.Request(url, headers={"User-Agent": UA})
    try:
        with urllib.request.urlopen(r, timeout=30) as resp:
            hits = json.loads(resp.read())
        time.sleep(1)  # respect Nominatim's ~1 req/sec policy
        if hits:
            ll = (str(hits[0]["lat"]), str(hits[0]["lon"]))
            return ll + (hits[0].get("boundingbox"),) if with_bbox else ll
    except Exception as e:
        print(f"  (geocoding failed: {e})", file=sys.stderr)
    return None


def grid_cells(bbox, n):
    """Split a Nominatim bbox [south, north, west, east] into n x n cells -> [{lat, lon, radius, zoom}].
    radius (m) = 1.15 x half the cell diagonal (equirectangular), at least 1000.
    zoom = the Maps zoom whose ~1000px viewport is about as wide as the cell (11..17), so each
    search is launched looking at its own tile instead of a city-wide or block-wide view."""
    s, nn, w, e = (float(x) for x in bbox)
    dlat, dlon = (nn - s) / n, (e - w) / n
    m_per_deg = 6371000 * math.pi / 180
    cells = []
    for i in range(n):
        for j in range(n):
            lat, lon = s + (i + 0.5) * dlat, w + (j + 0.5) * dlon
            cos = math.cos(math.radians(lat))
            width = max(1.0, dlon * m_per_deg * cos)
            diag = math.hypot(dlat * m_per_deg, width)
            zoom = round(math.log2(156543.03 * cos * 1000 / width))  # 156543 m/px at zoom 0
            cells.append({"lat": f"{lat:.6f}", "lon": f"{lon:.6f}",
                          "radius": max(1000, math.ceil(1.15 * diag / 2)),
                          "zoom": min(17, max(11, zoom))})
    return cells


def in_bbox(row, bbox, margin=0.10):
    """True if the row's latitude/longitude sit inside bbox [south, north, west, east] grown by
    `margin` of its size on each side. Rows with no usable coordinates are kept (True)."""
    try:
        lat, lon = float(row.get("latitude")), float(row.get("longitude"))
    except (TypeError, ValueError):
        return True
    s, n, w, e = (float(x) for x in bbox)
    my, mx = (n - s) * margin, (e - w) * margin
    return s - my <= lat <= n + my and w - mx <= lon <= e + mx


def dedupe(rows):
    """Keep the first row per place_id, else cid, else link, else title+address."""
    seen, out = set(), []
    for r in rows:
        key = next((f"{k}:{r[k]}" for k in ("place_id", "cid", "link") if r.get(k)),
                   f"ta:{r.get('title', '')}|{r.get('address', '')}")
        if key not in seen:
            seen.add(key); out.append(r)
    return out


def run_job(body):
    """Create one engine job, poll it, download it -> (job_id, raw CSV rows). Raises RuntimeError."""
    try:
        _, raw = req("POST", "/api/v1/jobs", body)
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"✗ Create failed: HTTP {e.code} — {e.read().decode()[:200]}")
    job_id = json.loads(raw).get("id")
    if not job_id:
        raise RuntimeError("✗ No job id returned.")
    print(f"  job id: {job_id}")

    print("▶ Polling…")
    status = None
    for i in range(120):
        _, raw = req("GET", f"/api/v1/jobs/{job_id}")
        status = json.loads(raw).get("Status")
        print(f"\r  status: {str(status):<10} (attempt {i + 1})", end="", flush=True)
        if status == "ok":
            print(); break
        if status == "failed":
            raise RuntimeError("\n✗ Job failed. If this keeps happening you may be rate-limited — wait or add proxies.")
        time.sleep(8)
    else:
        raise RuntimeError("\n✗ Timed out.")

    _, raw = req("GET", f"/api/v1/jobs/{job_id}/download")
    return job_id, list(csv.DictReader(io.StringIO(raw.decode("utf-8", "replace"))))
