#!/usr/bin/env python3
"""Builds data.json for the AFC Winkfield player dashboard.

Sheet mode (used when both links are set) reads two Google Sheet tabs published as CSV:
  SHEET_APPEARANCES_URL, SHEET_STATS_URL
Otherwise it reads two Baserow tables through the API (standard library only).
  BASEROW_TOKEN       database token (Baserow > Settings > Database tokens)
  TABLE_APPEARANCES   numeric id of the Appearances table
  TABLE_SEASON_STATS  numeric id of the Season Stats table
  BASEROW_API         optional, default https://api.baserow.io

Optional third table for the club page (writes club.json):
  TABLE_MATCHES (Baserow)  or  SHEET_MATCHES_URL (Google Sheet CSV link)
  Fields: Match ID, Match Date, Opponent, Score, Teams, Season, Managers, Result, Clean Sheet, Squad Size

Test mode reads the Excel workbook instead:  python build_data.py --xlsx stats.xlsx

Fields used
  Appearances : Name, Match ID, Season, Team, Manager, Opponent,
                Win (Number), Draw, Loss, Clean Sheet
  Season Stats: Name, Season, Goals, Assists, Yellow Cards, Red Cards
"""
import json, os, re, sys, time, datetime
from datetime import date
from collections import defaultdict, Counter
from urllib.request import Request, urlopen
from urllib.error import HTTPError

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.json")


# ---------- loading ----------
def fetch_table(api, token, table_id):
    rows, url = [], f"{api}/api/database/rows/table/{table_id}/?user_field_names=true&size=200"
    while url:
        for attempt in range(5):
            try:
                req = Request(url, headers={"Authorization": f"Token {token}", "Accept": "application/json", "User-Agent": "Mozilla/5.0 (compatible; afc-winkfield-dashboard/1.0)"})
                with urlopen(req, timeout=60) as r:
                    page = json.load(r)
                break
            except HTTPError as e:
                if e.code == 429 and attempt < 4:
                    time.sleep(2 ** attempt)
                    continue
                body = e.read().decode("utf-8", "replace")[:300]
                raise SystemExit(f"Baserow returned {e.code} for table {table_id}. Check the token, its read permission and the table id. Response: {body}")
        rows += page["results"]
        url = page.get("next")
    return rows


def load_csv(url, label):
    import csv, io
    try:
        req = Request(url, headers={"User-Agent": "Mozilla/5.0 (compatible; afc-winkfield-dashboard/1.0)"})
        with urlopen(req, timeout=120) as r:
            text = r.read().decode("utf-8-sig")
    except Exception as e:
        raise SystemExit(f"Could not download the {label} CSV link: {e}")
    if text.lstrip().lower().startswith("<!doctype html") or "<html" in text[:200].lower():
        raise SystemExit(f"The {label} link returned a web page, not CSV. Publish that tab to the web as CSV and use that link.")
    rows = list(csv.DictReader(io.StringIO(text)))
    return [{(k or "").strip(): v for k, v in r.items()} for r in rows]


def load_xlsx(path):
    import pandas as pd
    x = pd.ExcelFile(path)
    matches = x.parse("Matches").to_dict("records")
    # In this workbook, dates the Sheet read as real dates have day and month swapped
    # (14/9/2019 stayed text, 7/9/2019 became 9 July). Undo that here, for the test path only.
    for r in matches:
        v = r.get("Match Date")
        if hasattr(v, "year") and hasattr(v, "day") and v.day <= 12:
            r["Match Date"] = f"{v.year:04d}-{v.day:02d}-{v.month:02d}"
        elif hasattr(v, "year"):
            r["Match Date"] = f"{v.year:04d}-{v.month:02d}-{v.day:02d}"
    return (x.parse("Appearances").to_dict("records"),
            x.parse("Season Stats").to_dict("records"),
            matches)


# ---------- cleaning ----------
def flat(v):
    """Baserow link, lookup and select fields arrive as lists or dicts. Reduce to text."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, (list, tuple)):
        return ", ".join(x for x in (flat(i) for i in v) if x)
    if isinstance(v, dict):
        return flat(v.get("value", v.get("name", "")))
    if isinstance(v, float) and v != v:
        return ""
    return str(v).strip()


def num(v):
    s = flat(v).split(",")[0].strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def check_fields(rows, needed, label):
    if not rows:
        raise SystemExit(f"The {label} table returned no rows. Check the table id and that the token has read permission on it.")
    have = set(rows[0].keys())
    missing = [f for f in needed if f not in have]
    if missing:
        raise SystemExit(f"{label}: these fields were not found: {missing}. Fields that do exist: {sorted(have)}")


# ---------- build ----------
def build(app_rows, stat_rows):
    apps = []
    for r in app_rows:
        name = flat(r.get("Name"))
        if not name:
            continue
        apps.append(dict(
            n=name, m=flat(r.get("Match ID")), s=flat(r.get("Season")), t=flat(r.get("Team")),
            mg=flat(r.get("Manager")), op=flat(r.get("Opponent")),
            w=int(num(r.get("Win (Number)"))), d=int(num(r.get("Draw"))),
            l=int(num(r.get("Loss"))), cs=int(num(r.get("Clean Sheet")))))

    stats = defaultdict(lambda: dict(g=0, a=0, y=0, r=0))
    for r in stat_rows:
        name, season = flat(r.get("Name")), flat(r.get("Season"))
        if not name:
            continue
        k = stats[(name, season)]
        k["g"] += num(r.get("Goals")); k["a"] += num(r.get("Assists"))
        k["y"] += num(r.get("Yellow Cards")); k["r"] += num(r.get("Red Cards"))

    seasons = sorted({a["s"] for a in apps if a["s"]})
    by_player = defaultdict(list)
    roster, match_season = defaultdict(set), {}
    for a in apps:
        by_player[a["n"]].append(a)
        roster[a["m"]].add(a["n"])
        match_season[a["m"]] = a["s"]

    # teammate counts: player -> key -> Counter(teammate)
    tm = defaultdict(lambda: defaultdict(Counter))
    for m, people in roster.items():
        s = match_season[m]
        for p in people:
            for q in people:
                if p != q:
                    tm[p]["all"][q] += 1
                    tm[p][s][q] += 1

    def top(counter, n):
        return [[k, v] for k, v in sorted(counter.items(), key=lambda kv: (-kv[1], kv[0]))[:n]]

    players = {}
    for name, rows in by_player.items():
        blocks = {}
        for key in ["all"] + seasons:
            sel = rows if key == "all" else [r for r in rows if r["s"] == key]
            if not sel:
                continue
            if key == "all":
                gs = [stats[k] for k in stats if k[0] == name]
            else:
                gs = [stats[(name, key)]] if (name, key) in stats else []
            teams = Counter(r["t"] for r in sel if r["t"])
            opp = Counter(r["op"] for r in sel if r["op"])
            mgr = Counter(r["mg"] for r in sel if r["mg"])
            blocks[key] = dict(
                a=len(sel), g=int(sum(x["g"] for x in gs)), **{"as": int(sum(x["a"] for x in gs))},
                y=int(sum(x["y"] for x in gs)), r=int(sum(x["r"] for x in gs)),
                cs=sum(r["cs"] for r in sel), w=sum(r["w"] for r in sel),
                d=sum(r["d"] for r in sel), l=sum(r["l"] for r in sel),
                t=dict(teams), tm=top(tm[name][key], 5), mg=top(mgr, 5),
                op=top(opp, 10), no=len(opp))
        players[name] = blocks

    # appearance rank within each filter (ties share a rank)
    for key in ["all"] + seasons:
        counts = sorted((b[key]["a"] for b in players.values() if key in b), reverse=True)
        for b in players.values():
            if key in b:
                b[key]["rk"] = 1 + sum(1 for c in counts if c > b[key]["a"])

    return dict(seasons=seasons, players=dict(sorted(players.items())))


# ---------- club page data ----------
def season_window(label):
    m = re.match(r"(\d\d)/(\d\d)", label or "")
    if not m:
        return None
    y = 2000 + int(m.group(1))
    return date(y, 8, 1), date(y + 1, 6, 30)


def parse_date(raw, season):
    """Returns (iso date, inside). Reads 2019-09-14, 14/9/2019 and 14/09/19 (day first)."""
    t = flat(raw)
    d = None
    try:
        d = date.fromisoformat(t[:10])
    except ValueError:
        m = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4}|\d{2})", t)
        if m:
            y = int(m.group(3))
            try:
                d = date(y + 2000 if y < 100 else y, int(m.group(2)), int(m.group(1)))
            except ValueError:
                d = None
    if d is None:
        return "", False
    w = season_window(season)
    return d.isoformat(), (not w) or (w[0] <= d <= w[1])


def club_score(score, result):
    """Score fields are not always written with our goals first, so Result decides the order."""
    m = re.match(r"\s*(\d+)\s*-\s*(\d+)", flat(score))
    if not m:
        return None
    a, b = int(m.group(1)), int(m.group(2))
    hi, lo = max(a, b), min(a, b)
    if result == "W":
        return hi, lo
    if result == "L":
        return lo, hi
    return a, b


def venue_of(score, result):
    """Home score is written first. Wins and losses show the venue; draws and shoot-outs cannot."""
    m = re.match(r"\s*(\d+)\s*-\s*(\d+)", flat(score))
    if not m or result not in ("W", "L"):
        return ""
    a, b = int(m.group(1)), int(m.group(2))
    if a == b:
        return ""
    ours_first = (a > b) if result == "W" else (a < b)
    return "H" if ours_first else "A"


def build_club(match_rows):
    out, outside, skipped = [], 0, 0
    for r in match_rows:
        mid = flat(r.get("Match ID"))
        res = flat(r.get("Result")).upper()[:1]
        season, team = flat(r.get("Season")), flat(r.get("Teams"))
        sc = club_score(r.get("Score"), res)
        if not mid or res not in ("W", "D", "L") or not sc or not season:
            skipped += 1
            continue
        d, inside = parse_date(r.get("Match Date"), season)
        outside += not inside
        out.append([int(num(mid)), d, season, team, flat(r.get("Opponent")), sc[0], sc[1], res,
                    int(num(r.get("Clean Sheet"))), int(num(r.get("Squad Size"))), flat(r.get("Managers")),
                    venue_of(r.get("Score"), res)])
    out.sort(key=lambda m: (m[1], m[0]))
    print(f"Matches: {len(out)} kept, {skipped} skipped, {outside} dated outside their season (check these in Baserow).")
    return dict(
        fields=["id", "date", "season", "team", "opp", "gf", "ga", "res", "cs", "squad", "mgr", "venue"],
        matches=out, datesOutsideSeason=outside)


def build_managers(app_rows):
    """One row per manager, player, season and team. Joint managers each get the credit."""
    agg = {}
    for r in app_rows:
        name, mg = flat(r.get("Name")), flat(r.get("Manager"))
        if not name or not mg:
            continue
        season, team = flat(r.get("Season")), flat(r.get("Team"))
        vals = [int(num(r.get("Win (Number)"))), int(num(r.get("Draw"))), int(num(r.get("Loss"))), int(num(r.get("Clean Sheet")))]
        for m in [x.strip() for x in mg.split(",") if x.strip()]:
            a = agg.setdefault((m, name, season, team), [0, 0, 0, 0, 0])
            a[0] += 1
            for i, v in enumerate(vals):
                a[i + 1] += v
    rows = [[*k, *v] for k, v in sorted(agg.items())]
    return dict(fields=["manager", "player", "season", "team", "a", "w", "d", "l", "cs"], rows=rows)


def write_json(path, data, label):
    if os.path.exists(path):
        try:
            old = json.load(open(path))
            if {k: v for k, v in old.items() if k != "updated"} == data:
                print(f"No change in {label}; leaving it as it is.")
                return
        except Exception:
            pass
    data = dict(data, updated=datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%MZ"))
    with open(path, "w") as f:
        json.dump(data, f, separators=(",", ":"), ensure_ascii=False)
    print(f"Wrote {label}")


def main():
    mt = None
    if "--xlsx" in sys.argv:
        app, st, mt = load_xlsx(sys.argv[sys.argv.index("--xlsx") + 1])
    elif os.environ.get("SHEET_APPEARANCES_URL") and os.environ.get("SHEET_STATS_URL"):
        app = load_csv(os.environ["SHEET_APPEARANCES_URL"], "Appearances")
        st = load_csv(os.environ["SHEET_STATS_URL"], "Season Stats")
        if os.environ.get("SHEET_MATCHES_URL"):
            mt = load_csv(os.environ["SHEET_MATCHES_URL"], "Matches")
    else:
        token = os.environ.get("BASEROW_TOKEN")
        ta, ts = os.environ.get("TABLE_APPEARANCES"), os.environ.get("TABLE_SEASON_STATS")
        if not (token and ta and ts):
            raise SystemExit("Set BASEROW_TOKEN, TABLE_APPEARANCES and TABLE_SEASON_STATS.")
        api = os.environ.get("BASEROW_API", "https://api.baserow.io").rstrip("/")
        app, st = fetch_table(api, token, ta), fetch_table(api, token, ts)
        if os.environ.get("TABLE_MATCHES"):
            mt = fetch_table(api, token, os.environ["TABLE_MATCHES"])

    check_fields(app, ["Name", "Match ID", "Season", "Team", "Manager", "Opponent", "Win (Number)", "Draw", "Loss", "Clean Sheet"], "Appearances")
    check_fields(st, ["Name", "Season", "Goals", "Assists", "Yellow Cards", "Red Cards"], "Season Stats")
    data = build(app, st)
    if not data["players"]:
        raise SystemExit("No appearances found. Check the table ids and field names.")
    write_json(OUT, data, "data.json")
    write_json(os.path.join(os.path.dirname(OUT), "managers.json"), build_managers(app), "managers.json")

    if mt is not None:
        check_fields(mt, ["Match ID", "Match Date", "Opponent", "Score", "Teams", "Season", "Managers", "Result", "Clean Sheet", "Squad Size"], "Matches")
        club = build_club(mt)
        if not club["matches"]:
            raise SystemExit("No usable matches found. Check the Matches table fields.")
        write_json(os.path.join(os.path.dirname(OUT), "club.json"), club, "club.json")
    else:
        print("No Matches table set (TABLE_MATCHES), so club.json was not updated.")


if __name__ == "__main__":
    main()
