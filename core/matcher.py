"""Name matching: FCU unit <-> vendor row (by unit reference) and FCU space <-> vent space."""
import re
from difflib import SequenceMatcher

_CODE = re.compile(r"^&?\s*(\d+)")
_NOISE = re.compile(r"&?\d{1,3}(?:\s*[-,]\s*\d{1,3})*\s*[-]?\s*|\bBD\s*\d(?:&\d)?\b|\b[BG]F\b|\bGF\b|\bX\d+\b|\b\d+F\b", re.I)


def norm(s):
    s = re.sub(r"[^A-Z0-9]+", " ", (s or "").upper())
    return re.sub(r"\s+", " ", s).strip()


_SYN = {"STO": "STORE", "STOR": "STORE", "RM": "ROOM", "TOI": "TOILET", "WC": "TOILET", "ELE": "ELECTRICAL",
        "ELEC": "ELECTRICAL", "COR": "CORRIDOR", "CORR": "CORRIDOR", "MEP": "MEP", "EQ": "EQUIPMENT",
        "GAMES": "GAME", "ENTR": "ENTRANCE", "ENT": "ENTRANCE", "SERV": "SERVICE", "PR": "PUMP"}


def name_core(s):
    """Space name without tag codes ('&001-BD1-BF-HK STORE' -> 'HK STORE'), abbreviations expanded."""
    s = norm(_NOISE.sub(" ", (s or "").upper().replace("&", " &")))
    return " ".join(_SYN.get(t, t) for t in s.split())


def _ratio(a, b):
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    sm = SequenceMatcher(None, a, b)
    if sm.real_quick_ratio() < 0.5:
        return 0.0
    r = sm.ratio()
    ta, tb = set(a.split()), set(b.split())
    jac = len(ta & tb) / len(ta | tb) if ta | tb else 0.0
    return max(r * 0.7 + jac * 0.3, jac if jac >= 0.99 else 0.0)


def ref_score(a, b):
    """Score two unit references / zone names (codes kept: they identify the unit)."""
    na, nb = norm(a), norm(b)
    s = _ratio(na, nb)
    ca, cb = _CODE.match((a or "").strip()), _CODE.match((b or "").strip())
    if ca and cb and int(ca.group(1)) == int(cb.group(1)):
        s = min(1.0, s + 0.15)
    return s


def space_score(a, b):
    return _ratio(name_core(a), name_core(b))


def _code(k):
    m = _CODE.match((k or "").strip())
    return int(m.group(1)) if m else None


def match_vendor(units, vendor_rows, threshold=0.6, topn=5):
    """-> per unit {'best': idx|None, 'score': float, 'cands': [(idx, score)...]}.

    Exact (normalised) name matches are taken first; units without one are
    fuzzy matched. Best matches are assigned one-to-one, highest score first.
    """
    vkeys = [[(norm(k), _code(k)) for k in v.keys] for v in vendor_rows]
    exact = {}
    for vi, ks in enumerate(vkeys):
        for n, _ in ks:
            exact.setdefault(n, []).append(vi)
    cand = []
    for u in units:
        ukeys = [(norm(k), _code(k)) for k in dict.fromkeys([u.zone, u.tag]) if k]
        hits = {vi: 1.0 for n, _ in ukeys for vi in exact.get(n, ())}
        if hits:
            scored = list(hits.items())
        else:
            scored = []
            for vi, ks in enumerate(vkeys):
                best = 0.0
                for na, ca in ukeys:
                    for nb, cb in ks:
                        sc = _ratio(na, nb)
                        if ca is not None and ca == cb:
                            sc = min(1.0, sc + 0.15)
                        best = max(best, sc)
                scored.append((vi, best))
            scored.sort(key=lambda x: -x[1])
        cand.append(scored[:topn])
    pairs = sorted(((s, ui, vi) for ui, c in enumerate(cand) for vi, s in c if s >= threshold), reverse=True)
    used_u, used_v, assign = set(), set(), {}
    for s, ui, vi in pairs:
        if ui in used_u or vi in used_v:
            continue
        used_u.add(ui)
        used_v.add(vi)
        assign[ui] = (vi, s)
    return [
        {"best": assign[i][0] if i in assign else None,
         "score": assign[i][1] if i in assign else (cand[i][0][1] if cand[i] else 0.0),
         "cands": cand[i]}
        for i in range(len(units))
    ]


def _bld_file(fname):
    m = re.search(r"BLD?G[\s\-_]*(\d)", fname or "", re.I)
    return m.group(1) if m else None


def _bld_name(name):
    m = re.search(r"BD\s*(\d)(?:\s*&\s*(\d))?", name or "", re.I)
    return {g for g in m.groups() if g} if m else set()


def _area_sim(a, b):
    if not a or not b:
        return 0.0
    return 1.0 - min(1.0, abs(a - b) / max(a, b))


def match_vent(units, vent_spaces, threshold=0.4, topn=8):
    """Match every FCU space to its closest ventilation space.

    Score = name similarity (+ building tag agreement, + area similarity as
    tie-breaker). The best candidate at or above `threshold` is always used;
    the reviewer can change it. Unit values later sum the distinct matches.

    -> per unit {'hits': [(space_name, vent_idx, score)], 'cands': [(idx, score)], 'score'}
    """
    cores = [name_core(v.name) for v in vent_spaces]
    bld = [_bld_file(v.file) for v in vent_spaces]
    out = []
    for u in units:
        spaces = [(s[0], s[1]) for s in u.spaces] or [(u.zone, u.area)]
        hits, allc, seen = [], {}, set()
        for nm, area in spaces:
            core, nb = name_core(nm), _bld_name(nm)
            scored = []
            for vi, c in enumerate(cores):
                r = _ratio(core, c)
                if r <= 0:
                    continue
                sc = r + 0.10 * _area_sim(area, vent_spaces[vi].area)
                if nb and bld[vi]:
                    sc += 0.1 if bld[vi] in nb else -0.1
                scored.append((vi, sc))
            scored.sort(key=lambda x: -x[1])
            for vi, sc in scored[:topn]:
                allc[vi] = max(sc, allc.get(vi, 0))
            if scored and scored[0][1] >= threshold and scored[0][0] not in seen:
                seen.add(scored[0][0])
                hits.append((nm, scored[0][0], min(scored[0][1], 1.0)))
        cands = sorted(allc.items(), key=lambda x: -x[1])[:topn]
        out.append({"hits": hits, "amb": [], "cands": cands,
                    "score": min((h[2] for h in hits), default=0.0)})
    return out
