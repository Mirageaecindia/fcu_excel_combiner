"""Name matching: FCU unit <-> vendor row (by unit reference) and FCU space <-> vent space."""
import re
from difflib import SequenceMatcher

_CODE = re.compile(r"^&?\s*(\d+)")
_NOISE = re.compile(r"&?\d{1,3}(?:\s*[-,]\s*\d{1,3})*\s*[-]?\s*|\bBD\s*\d(?:&\d)?\b|\b[BG]F\b|\bGF\b|\bX\d+\b|\b\d+F\b", re.I)


def norm(s):
    s = re.sub(r"[^A-Z0-9]+", " ", (s or "").upper())
    return re.sub(r"\s+", " ", s).strip()


def name_core(s):
    """Space name without tag codes ('&001-BD1-BF-HK STORE' -> 'HK STORE')."""
    s = _NOISE.sub(" ", (s or "").upper().replace("&", " &"))
    return norm(s)


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


def match_vent(units, vent_spaces, threshold=0.75, topn=6):
    """Match each FCU space to a vent space; unit result sums people / fresh air.

    A match is only taken automatically when it is unambiguous: exactly one vent
    space has the best score and no other FCU space claims the same vent space.
    Ambiguous spaces are listed so the reviewer can pick the right one.

    -> per unit {'hits': [(space_name, vent_idx, score)], 'amb': [space_name],
                 'cands': [(idx, score)], 'score'}
    """
    cores = [name_core(v.name) for v in vent_spaces]
    bld = [_bld_file(v.file) for v in vent_spaces]
    claims = {}
    raw = []
    for ui, u in enumerate(units):
        names = [s[0] for s in u.spaces] or [u.zone]
        per = []
        for nm in names:
            core, nb = name_core(nm), _bld_name(nm)
            scored = sorted(
                ((vi, _ratio(core, c) + (0.02 if bld[vi] and bld[vi] in nb else 0.0)) for vi, c in enumerate(cores)),
                key=lambda x: -x[1])
            top = scored[:topn]
            ok = [x for x in scored if x[1] >= threshold]
            best = ok[0] if ok else None
            tied = best is not None and len(ok) > 1 and abs(ok[1][1] - best[1]) < 1e-6
            if best and not tied:
                claims.setdefault(best[0], []).append(ui)
            per.append((nm, best, tied, top))
        raw.append(per)
    out = []
    for ui, per in enumerate(raw):
        hits, amb, allc = [], [], {}
        for nm, best, tied, top in per:
            for vi, s in top:
                allc[vi] = max(s, allc.get(vi, 0))
            if best is None:
                continue
            if tied or len(claims.get(best[0], [])) > 1:
                amb.append(nm)
            else:
                hits.append((nm, best[0], min(best[1], 1.0)))
        cands = sorted(allc.items(), key=lambda x: -x[1])[:topn]
        out.append({"hits": hits, "amb": amb, "cands": cands,
                    "score": min((h[2] for h in hits), default=0.0)})
    return out
