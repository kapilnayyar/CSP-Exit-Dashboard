"""Name collision resolver — deterministic, no manual maintenance.

Problem
-------
Some CSPs share a name in Supabase (e.g., two "Shree Shyam Broadband", two
"Riddhi Enterprises"). Sheets like Migration Data and Main sheet key rows by
name only. Without a rule, we would double-count. Historically we maintained
a hand-edited `ATTRIBUTION_OVERRIDE` map. Kapil no longer wants to touch it.

Rule
----
For each set of partners that share a lowercased name:

    owner  = the ONE partner with the "highest" tuple:
             (u1_count OR u2_count,  # more customers = more likely owns the sheet rows
              state_index,           # tiebreak: further-along wins
              exit_started_at,       # tiebreak: more recent wins
              -partner_code)         # final tiebreak: lower code wins

    losers = every other partner in the same name group

Rationale
---------
The primary signal is `u1_count` / `u2_count` from Supabase's
`partner_details_extended` view. The partner with more customers is the one
whose Migration Data / Main sheet rows we've been reading; the low-customer
partner is a shadow that would double-count if we didn't force it to 0. This
reproduces every hand-maintained ATTRIBUTION_OVERRIDE decision automatically.
Falls back to state / exit date / code for the (rare) ties.

Public API
----------
resolve_collisions(partners) returns:
    {
      "name_to_owner_code": {lowercased_name -> partner_code_str},
      "losers": set(partner_code_str),
      "collisions": [ (name, owner_code, [loser_codes]) , ... ]
    }
"""
from collections import defaultdict
from datetime import datetime


# S6 is furthest along, S0 is the earliest. Higher index = wins.
STATE_INDEX = {"S0": 0, "S1": 1, "S2": 2, "S3": 3, "S4": 4, "S5": 5, "S6": 6}


def _parse_dt(s):
    """Return a comparable datetime; missing/invalid → epoch."""
    if not s:
        return datetime(1970, 1, 1)
    try:
        return datetime.fromisoformat(str(s).replace("Z", "+00:00")).replace(tzinfo=None)
    except Exception:
        return datetime(1970, 1, 1)


def _sort_key(p):
    """Higher tuple = winner. Used with max().
    Primary signal: total customers (u1_count + u2_count) — the partner with
    more customers is the one whose sheet rows we've been counting.
    """
    total_customers = int(p.get("u1_count") or 0) + int(p.get("u2_count") or 0)
    return (
        total_customers,
        STATE_INDEX.get(str(p.get("current_state") or ""), 0),
        _parse_dt(p.get("exit_started_at")),
        -int(p.get("partner_code") or 0),
    )


def resolve_collisions(partners, verbose=True):
    """Given a list of partner dicts (with name, partner_code, current_state,
    exit_started_at), return the collision map.

    Parameters
    ----------
    partners : list of dicts
        Each dict must have: name, partner_code, current_state, exit_started_at,
        u1_count, u2_count (both integers; missing → 0).
    verbose : bool
        If True, print a summary line per resolved collision.

    Returns
    -------
    dict with keys:
        name_to_owner_code : {name.lower(): str(partner_code)}
            Every unique lowercased name maps to the owner's partner_code
            (as string). For non-collision names, this is just their own code.
        losers : set of str(partner_code)
            Codes that lost their collision and MUST be forced to 0 for any
            name-based sheet lookup.
        collisions : list of (name, owner_code_str, [loser_code_str])
            For logging / caption in the dashboard.
    """
    if not partners:
        return {"name_to_owner_code": {}, "losers": set(), "collisions": []}

    by_name = defaultdict(list)
    for p in partners:
        nm = str(p.get("name") or "").strip().lower()
        if not nm:
            continue
        by_name[nm].append(p)

    name_to_owner_code = {}
    losers = set()
    collisions = []

    for name, group in by_name.items():
        if len(group) == 1:
            name_to_owner_code[name] = str(group[0]["partner_code"])
            continue

        # Collision — pick winner deterministically
        winner = max(group, key=_sort_key)
        winner_code = str(winner["partner_code"])
        name_to_owner_code[name] = winner_code

        loser_list = []
        for p in group:
            code = str(p["partner_code"])
            if code != winner_code:
                losers.add(code)
                loser_list.append(code)

        collisions.append((name, winner_code, loser_list))
        if verbose:
            print(
                f"  collision resolved: {name!r} "
                f"-> owner {winner_code} ({winner.get('current_state')}), "
                f"losers {loser_list}"
            )

    return {
        "name_to_owner_code": name_to_owner_code,
        "losers": losers,
        "collisions": collisions,
    }


if __name__ == "__main__":
    # Quick smoke-test against live Supabase
    import os
    import requests
    from dotenv import load_dotenv

    load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))
    SB = os.getenv("SUPABASE_URL")
    KEY = os.getenv("SUPABASE_ANON_KEY")
    hdr = {"apikey": KEY, "Authorization": f"Bearer {KEY}", "Accept": "application/json"}
    r = requests.get(
        f"{SB}/rest/v1/partners",
        params={"select": "id,name,partner_code,current_state,exit_started_at"},
        headers=hdr,
        timeout=15,
    ).json()
    # Attach u1_count / u2_count from partner_details_extended
    r2 = requests.get(
        f"{SB}/rest/v1/partner_details_extended",
        params={"select": "id,u1_count,u2_count"},
        headers=hdr,
        timeout=15,
    ).json()
    de = {row["id"]: row for row in r2}
    for p in r:
        d = de.get(p["id"], {})
        p["u1_count"] = d.get("u1_count") or 0
        p["u2_count"] = d.get("u2_count") or 0
    print(f"Loaded {len(r)} partners from Supabase")
    out = resolve_collisions(r, verbose=True)
    print(f"\nSummary: {len(out['collisions'])} collision(s), "
          f"{len(out['losers'])} loser code(s) forced to 0.")
