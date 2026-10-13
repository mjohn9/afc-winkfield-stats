#!/usr/bin/env python3
"""Builds data.json for the AFC Winkfield player dashboard.

Default mode reads two Baserow tables through the API (standard library only).
  BASEROW_TOKEN       database token (Baserow > Settings > Database tokens)
  TABLE_APPEARANCES   numeric id of the Appearances table
  TABLE_SEASON_STATS  numeric id of the Season Stats table
  BASEROW_API         optional, default https://api.baserow.io

Test mode reads the Excel workbook instead:  python build_data.py --xlsx stats.xlsx

Fields used
  Appearances : Name, Match ID, Season, Team, Manager, Opponent,
                Win (Number), Draw, Loss, Clean Sheet
  Season Stats: Name, Season, Goals, Assists, Yellow Cards, Red Cards
"""
import json, os, sys, time, datetime
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
                req = Request(url, headers={"Authorization": f"Token {token}"})
                with urlopen(req, timeout=60) as r:
                    page = json.load(r)
                break
            except HTTPError as e:
                if e.code == 429 and attempt < 4:
                    time.sleep(2 ** attempt)
                    continue
                raise SystemExit(f"Baserow returned {e.code} for table {table_id}. Check the token and table id.")
        rows += page["results"]
        url = page.get("next")
    return rows


def load_xlsx(path):
    import pandas as pd
    x = pd.ExcelFile(path)
    return (x.parse("Appearances").to_dict("records"),
            x.parse("Season Stats").to_dict("records"))


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


def main():
    if "--xlsx" in sys.argv:
        app, st = load_xlsx(sys.argv[sys.argv.index("--xlsx") + 1])
    else:
        token = os.environ.get("BASEROW_TOKEN")
        ta, ts = os.environ.get("TABLE_APPEARANCES"), os.environ.get("TABLE_SEASON_STATS")
        if not (token and ta and ts):
            raise SystemExit("Set BASEROW_TOKEN, TABLE_APPEARANCES and TABLE_SEASON_STATS.")
        api = os.environ.get("BASEROW_API", "https://api.baserow.io").rstrip("/")
        app, st = fetch_table(api, token, ta), fetch_table(api, token, ts)

    data = build(app, st)
    if not data["players"]:
        raise SystemExit("No appearances found. Check the table ids and field names.")

    if os.path.exists(OUT):
        try:
            old = json.load(open(OUT))
            if {k: v for k, v in old.items() if k != "updated"} == data:
                print("No change in data; leaving data.json as it is.")
                return
        except Exception:
            pass
    data["updated"] = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%MZ")
    with open(OUT, "w") as f:
        json.dump(data, f, separators=(",", ":"), ensure_ascii=False)
    print(f"Wrote {OUT}: {len(data['players'])} players, seasons {data['seasons']}")


if __name__ == "__main__":
    main()
