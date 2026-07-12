"""
cross_check_thesis.py
======================
Cross-checks every numeric table cell (and a best-effort scan of inline
narrative numbers) in Infante_Thesis_Consolidated_Revision_v2.docx against
the analysis pipeline's output CSVs (outputs/) and numbers_manifest.csv.

READ-ONLY: never modifies the docx. Writes discrepancy_report.csv at the
repo root.

Usage: python cross_check_thesis.py

2026-07-12: discrepancy adjudication
-------------------------------------
Classification is now rounding-convention-aware: a thesis value is only
flagged if it does NOT equal its source rounded half-up at the thesis's
OWN displayed decimal precision (see check_cell()). Values that match
under that rule are a silent PASS, not a ROUNDING line -- the previous
version flagged several cells (e.g. 31.41 vs a stored 31.405) purely
because it rounded with Python's banker's-rounding-adjacent round()
instead of round-half-up, or compared against the wrong decimal place.
"""
import re
import io
import sys
import json
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

import docx
import pandas as pd
from docx.document import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

REPO = Path(r"C:\Users\lucho\OneDrive\Escritorio\Thesis Master")
DOCX_PATH = REPO / "Infante_Thesis_Consolidated_Revision_v3.docx"
OUT = REPO / "outputs"
MANIFEST_PATH = REPO / "numbers_manifest.csv"
REPORT_PATH = REPO / "discrepancy_report.csv"

discrepancies = []  # list of dicts: location_in_docx, thesis_value, source_value, source_file, severity


def flag(location, thesis_value, source_value, source_file, severity):
    discrepancies.append(dict(location_in_docx=location, thesis_value=thesis_value,
                               source_value=source_value, source_file=source_file,
                               severity=severity))


def num(s):
    """Parse a thesis-formatted number ('−67%', '$1.30 bn', '1,957', '0.052×') to float, or None."""
    if s is None:
        return None
    s = str(s)
    s = s.replace('−', '-').replace(',', '').replace('×', '').replace('%', '')
    s = re.sub(r'\$|bn|bp|m\b', '', s).strip()
    m = re.search(r'-?\d+\.?\d*', s)
    return float(m.group()) if m else None


def close(a, b, tol):
    if a is None or b is None:
        return False
    return abs(a - b) <= tol


# ------------------------------------------------------------------
# 2026-07-12: discrepancy adjudication -- rounding-convention-aware check
# ------------------------------------------------------------------

def displayed_decimals(raw_str):
    """Count the decimal digits actually shown in the thesis's own text."""
    m = re.search(r'\.(\d+)', str(raw_str))
    return len(m.group(1)) if m else 0


def round_half_up(value, decimals):
    """Round-half-up (ties away from zero), via Decimal for exactness --
    avoids Python's round() banker's-rounding-plus-float-imprecision
    surprises (e.g. round(31.405, 2) == 31.4, not 31.41)."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except TypeError:
        pass
    try:
        q = Decimal(1).scaleb(-decimals)
        d = Decimal(str(value))
        if value < 0:
            return -float((-d).quantize(q, rounding=ROUND_HALF_UP))
        return float(d.quantize(q, rounding=ROUND_HALF_UP))
    except InvalidOperation:
        return None


def check_cell(location, thesis_raw, source_value, source_file):
    """Compare a thesis-displayed number against source_value using
    round-half-up at the thesis's OWN displayed precision. A value equal
    to the correctly-rounded source is a silent PASS (nothing flagged);
    only genuine deviations from that rounding are reported, split into
    ROUNDING (off by about one unit at the last displayed digit -- a
    tie-breaking or double-rounding artifact) vs MISMATCH (everything
    else). Returns the severity string, or None if it passed."""
    docv = num(thesis_raw)
    if docv is None or source_value is None:
        return None
    try:
        if pd.isna(source_value):
            return None
    except TypeError:
        pass
    dp = displayed_decimals(thesis_raw)
    expected = round_half_up(source_value, dp)
    if expected is not None and abs(docv - expected) < 1e-9:
        return None  # PASS
    ulp = 10 ** -dp
    sev = 'ROUNDING' if abs(docv - source_value) <= ulp * 1.01 else 'MISMATCH'
    flag(location, thesis_raw, round_half_up(source_value, dp + 1), source_file, sev)
    return sev


# ==============================================================================
# 1. Extract every table from the docx
# ==============================================================================

def iter_block_items(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    else:
        parent_elm = parent._element
    for child in parent_elm.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, parent)
        elif child.tag == qn('w:tbl'):
            yield Table(child, parent)


TABLE_RE = re.compile(r'\bTable\s+(\d+\.\d+|A\.\d+)\b')

d = docx.Document(str(DOCX_PATH))
blocks = list(iter_block_items(d))
items = []
for b in blocks:
    if isinstance(b, Paragraph):
        t = b.text.strip()
        if t:
            items.append(('para', b.style.name if b.style else '', t))
    else:
        items.append(('table', None, b))

n = len(items)
docx_tables = {}  # label -> list[list[str]] (rows incl. header)
# NOTE: in this document, a table's caption paragraph sits AFTER the table
# (confirmed by inspection), so resolve forward first and only fall back to
# a backward search if no caption follows (e.g. the very last table before a
# section break). Searching backward-first would silently attach each table
# to the PRECEDING table's caption -- a mislabeling bug caught during review.
for i, it in enumerate(items):
    if it[0] != 'table':
        continue
    label = None
    for j in range(i + 1, min(n, i + 6)):
        if items[j][0] == 'para':
            m = TABLE_RE.search(items[j][2])
            if m:
                label = m.group(1)
                break
        else:
            break
    if label is None:
        for j in range(i - 1, max(-1, i - 6), -1):
            if items[j][0] == 'para':
                m = TABLE_RE.search(items[j][2])
                if m:
                    label = m.group(1)
                    break
            else:
                break
    rows = [[c.text.strip().replace('\n', ' ') for c in row.cells] for row in it[2].rows]
    if label in docx_tables:
        print(f"WARNING: label {label!r} already assigned; this table (idx {i}) collides.")
    docx_tables[label] = rows

print(f"Extracted {len(docx_tables)} labelled tables from the docx: {sorted(docx_tables)}")


def row_dicts(label):
    """Return table rows as list of dicts keyed by (deduplicated) header."""
    rows = docx_tables[label]
    header = rows[0]
    # de-duplicate merged-cell artifacts (consecutive identical header text)
    seen = []
    keep_idx = []
    for i, h in enumerate(header):
        if i > 0 and h == header[i - 1]:
            continue
        seen.append(h)
        keep_idx.append(i)
    out = []
    for r in rows[1:]:
        out.append({seen[k]: r[idx] for k, idx in enumerate(keep_idx)})
    return out


# ==============================================================================
# 2. Load source CSVs
# ==============================================================================

def load(name):
    return pd.read_csv(OUT / name)


cb_ts = load('capital_bridge_timeseries.csv')
retained = load('retained_earnings_attribution.csv')
eps = load('rate_shock_episodes.csv')
mtm_grid = load('mtm_loss_grid.csv')
lcr_grid = load('lcr_stress_grid.csv')
lcr_eff = load('lcr_effective_coverage.csv')
cs_ts = load('combined_stress_timeseries.csv').drop_duplicates('Quarter').set_index('Quarter')
sp = load('btc_gold_shock_panel.csv').drop_duplicates('Quarter').set_index('Quarter')
joint_replays = load('joint_stress_replays.csv')
score = load('classification_scorecard.csv')
schemes = load('classification_weight_sensitivity.csv')
struct_events = load('structural_events_detected.csv')
descriptive = load('descriptive_table.csv')
rwa_table = load('rwa_table.csv')
usdp_rwa = load('usdp_rwa_table.csv')
share = load('stablecoin_tbill_share.csv')
bank_bench = load('bank_benchmark.csv')

manifest = pd.read_csv(MANIFEST_PATH) if MANIFEST_PATH.exists() else pd.DataFrame()

# ==============================================================================
# 3. Per-table comparisons
# ==============================================================================

# ---- Table 4.1: risk-weight mapping (qualitative; weights are the only numbers) ----
if '4.1' in docx_tables:
    expected_weights = {
        'US Treasury bills (direct or via CRF look-through)': 0,
        'Reverse repos, UST-collateralised (overnight and term)': 0,
        'Cash at commercial banks': 20,
        'MMF units, opaque (USDT; USDC sensitivity case)': 20,
        'Physical gold': 100,
        'Secured loans': 100,
        'Unclassified residual': 100,
        'Bitcoin': 1250,
    }
    for r in row_dicts('4.1'):
        asset = r.get('Asset line', '')
        if asset in expected_weights:
            docv = num(r.get('Weight'))
            exp = expected_weights[asset]
            if docv != exp:
                flag(f'Table 4.1, {asset}', r.get('Weight'), f'{exp}%',
                     'Methodology sheet (Stablecoin Reserve Panel)', 'MISMATCH')

# ---- Table 4.2: capital bridge (E, RWA, cov), all 14 quarters ----
if '4.2' in docx_tables:
    for r in row_dicts('4.2'):
        q = r['Quarter']
        for iss, e_key, rwa_key, cov_key in [('USDC', 'USDC E ($bn)', 'USDC RWA ($bn)', 'USDC cov.'),
                                              ('USDT', 'USDT E ($bn)', 'USDT RWA ($bn)', 'USDT cov.')]:
            src = cb_ts[(cb_ts.Issuer == iss) & (cb_ts.Quarter == q)]
            if src.empty:
                flag(f'Table 4.2, {q}, {iss}', r.get(e_key), None,
                     'capital_bridge_timeseries.csv', 'NOT_FOUND_IN_OUTPUTS')
                continue
            src = src.iloc[0]
            check_cell(f'Table 4.2, {q}, {e_key}', r.get(e_key),
                       src['Reserve_Equity_bn'], 'capital_bridge_timeseries.csv')
            check_cell(f'Table 4.2, {q}, {rwa_key}', r.get(rwa_key),
                       src['RWA_bn'], 'capital_bridge_timeseries.csv')
            check_cell(f'Table 4.2, {q}, {cov_key}', r.get(cov_key),
                       src['Coverage_x'], 'capital_bridge_timeseries.csv')
        # USDP coverage
        src_p = cb_ts[(cb_ts.Issuer == 'USDP') & (cb_ts.Quarter == q)]
        if not src_p.empty:
            check_cell(f'Table 4.2, {q}, USDP cov.', r.get('USDP cov.'),
                       src_p.iloc[0]['Coverage_x'], 'capital_bridge_timeseries.csv')

# ---- Table 4.3: retained-earnings attribution ----
if '4.3' in docx_tables:
    rmap = {(r['Issuer'], r['Year']): r for r in retained.to_dict('records')}
    for r in row_dicts('4.3'):
        key = (r['Issuer'], r['Year'])
        if key not in rmap:
            continue
        src = rmap[key]
        check_cell(f'Table 4.3, {r["Issuer"]} {r["Year"]}, Retention', r.get('Retention'),
                   src['Retention_pct'], 'retained_earnings_attribution.csv')

# ---- Table 5.1: historical rate-shock episodes ----
if '5.1' in docx_tables:
    for r, (_, src) in zip(row_dicts('5.1'), eps.iterrows()):
        check_cell(f'Table 5.1, {r["Episode"]}, Peak rise', r.get('Peak rise (bp)'),
                   src['peak_bp'], 'rate_shock_episodes.csv')
        docm = num(r.get('Months ≥ +400bp'))
        if docm != src['n_months']:
            flag(f'Table 5.1, {r["Episode"]}, Months', r.get('Months ≥ +400bp'),
                 src['n_months'], 'rate_shock_episodes.csv', 'MISMATCH')

# ---- Table 5.2: duration stress grid corners ----
if '5.2' in docx_tables:
    for r in row_dicts('5.2'):
        iss = r['Issuer']
        mild = mtm_grid[(mtm_grid.Issuer == iss) & (mtm_grid.WAM_days == 30) & (mtm_grid.Shock_bp == 100)].iloc[0]
        worst = mtm_grid[(mtm_grid.Issuer == iss) & (mtm_grid.WAM_days == 180) & (mtm_grid.Shock_bp == 400)].iloc[0]
        check_cell(f'Table 5.2, {iss}, Mildest cell', r.get('Mildest cell (30d, +100bp)'),
                   mild['MtM_Loss_pct_of_assets'], 'mtm_loss_grid.csv')
        check_cell(f'Table 5.2, {iss}, Worst cell %', r.get('Worst cell (180d, +400bp)'),
                   worst['MtM_Loss_pct_of_assets'], 'mtm_loss_grid.csv')

# ---- Table 6.1: LCR stress grid, Q4 2025 ----
if '6.1' in docx_tables:
    treat_map = {('USDC', 'look-through'): 'lookthrough', ('USDC', 'opaque (sens.)'): 'opaque',
                 ('USDT', 'single'): 'lookthrough', ('USDP', 'single'): 'single'}
    for r in row_dicts('6.1'):
        key = (r['Issuer'], r['Treatment'])
        treat = treat_map.get(key)
        if treat is None:
            continue
        for outflow, col in [(20, '20% run'), (40, '40% run'), (60, '60% run')]:
            row = lcr_grid[(lcr_grid.Issuer == r['Issuer']) & (lcr_grid.Outflow_pct == outflow) &
                           (lcr_grid.Treatment == treat)]
            if row.empty:
                continue
            check_cell(f'Table 6.1, {r["Issuer"]}/{r["Treatment"]}, {col}', r.get(col),
                       row.iloc[0]['LCR_pct'], 'lcr_stress_grid.csv')

# ---- Table 6.2: effective LCR (40% run) ----
if '6.2' in docx_tables:
    l40 = lcr_eff[lcr_eff['Outflow_pct'] == 40].set_index('Quarter')
    for r in row_dicts('6.2'):
        q = r['Quarter']
        if q not in l40.index:
            continue
        src = l40.loc[q]
        for docc, srccol in [('USDC plain', 'USDC_LCR_plain'), ('USDC eff.', 'USDC_LCR_effective'),
                              ('USDC 5y-cf.', 'USDC_LCR_effective_5y'), ('USDT plain', 'USDT_LCR_plain'),
                              ('USDT eff.', 'USDT_LCR_effective'), ('USDT 5y-cf.', 'USDT_LCR_effective_5y')]:
            check_cell(f'Table 6.2, {q}, {docc}', r.get(docc), src[srccol], 'lcr_effective_coverage.csv')

# ---- Table 6.3: combined-stress timeline (sleeve/buffer/break-even/replay) ----
# NOTE: Sleeve is a derived sum (BTC_Notional_bn + Gold_Notional_bn); there is
# no single committed cell for it. 2024-Q1 Break-even is handled specially in
# Task 2a (recomputed unrounded from panel inputs, not from the CSV's own
# already-1dp-rounded Breakeven_Combined_pct) rather than via check_cell here.
if '6.3' in docx_tables:
    for r in row_dicts('6.3'):
        q = r['Quarter']
        if q not in sp.index or q not in cs_ts.index:
            continue
        srow, crow = sp.loc[q], cs_ts.loc[q]
        sleeve_src = float(srow['BTC_Notional_bn']) + float(srow['Gold_Notional_bn'])
        check_cell(f'Table 6.3, {q}, Sleeve', r.get('Sleeve ($bn)'), sleeve_src,
                   'btc_gold_shock_panel.csv (derived: BTC_Notional_bn + Gold_Notional_bn)')
        check_cell(f'Table 6.3, {q}, Buffer', r.get('Buffer ($bn)'),
                   srow['Reserve_Equity_bn'], 'btc_gold_shock_panel.csv')
        be = srow['Breakeven_Combined_pct']
        if pd.notna(be) and q != '2024-Q1':  # 2024-Q1 handled in Task 2a
            check_cell(f'Table 6.3, {q}, Break-even shock', r.get('Break-even shock'),
                       -be, 'btc_gold_shock_panel.csv')
        check_cell(f'Table 6.3, {q}, Replay residual', r.get('Q2-2022 replay residual ($bn)'),
                   crow['Replay_2022Q2_Buffer_after_bn'], 'combined_stress_timeseries.csv')

# ---- Table 6.4: historical joint episodes replay ----
# NOTE: the same calendar quarter can appear in BOTH the "triple (BTC era)"
# and "dual (gold+rates)" families with different Loss/Buffer values (e.g.
# 2022-Q3), so the join key must be (Episode, family), not Episode alone --
# an earlier version of this script collapsed them via a plain dict keyed
# only by Episode and silently compared against the wrong family's row.
if '6.4' in docx_tables:
    def family_of(type_str):
        return 'dual' if 'dual' in type_str else 'triple'
    jmap = {(r['Episode'], family_of(r['Type'])): r for r in joint_replays.to_dict('records')}
    for r in row_dicts('6.4'):
        is_dual = '(gold+rates)' in r['Episode']
        ep = r['Episode'].split(' (')[0]
        key = (ep, 'dual' if is_dual else 'triple')
        if key not in jmap:
            flag(f'Table 6.4, {r["Episode"]}', r.get('Loss on Q4-2025 book ($bn)'), None,
                 'joint_stress_replays.csv', 'NOT_FOUND_IN_OUTPUTS')
            continue
        src = jmap[key]
        check_cell(f'Table 6.4, {r["Episode"]}, Loss', r.get('Loss on Q4-2025 book ($bn)'),
                   src['Loss_bn'], 'joint_stress_replays.csv')
        check_cell(f'Table 6.4, {r["Episode"]}, Residual buffer', r.get('Residual buffer ($bn)'),
                   src['Buffer_after_bn'], 'joint_stress_replays.csv')

# ---- Table 7.1: raw classification axis values ----
# NOTE: this table's header row has "LCR at 40% run (mean)" merged across two
# grid columns, but the BODY rows' merge boundary is shifted one column to
# the left (the MtM cell is what's actually duplicated in the body, not
# LCR) -- a header/body merged-cell misalignment specific to this table.
# Generic header-based de-duplication silently mismatches columns here, so
# this table is handled with hardcoded positional indices instead.
if '7.1' in docx_tables:
    smap = {r['Issuer']: r for r in score.to_dict('records')}
    for raw in docx_tables['7.1'][1:]:
        issuer = raw[0]
        src = smap.get(issuer)
        if src is None:
            continue
        # raw = [Issuer, RWA, MtM, <duplicate-of-MtM artifact>, LCR, OffBasel, Disc]
        pairs = [(raw[1], 'raw_rwa_pct', 'RWA density (mean)'),
                 (raw[2], 'raw_mtm_pct', 'MtM adverse ceiling (mean)'),
                 (raw[4], 'raw_lcr_pct', 'LCR at 40% run (mean)'),
                 (raw[5], 'raw_risk_pct', 'Off-Basel share (mean)'),
                 (raw[6], 'raw_disc', 'Disc.')]
        for docraw, srccol, label in pairs:
            check_cell(f'Table 7.1, {issuer}, {label}', docraw, src[srccol],
                       'classification_scorecard.csv')

# ---- Table 7.2: sub-scores + blended score ----
if '7.2' in docx_tables:
    smap = {r['Issuer']: r for r in score.to_dict('records')}
    for r in row_dicts('7.2'):
        src = smap.get(r['Issuer'])
        if src is None:
            continue
        for doccol, srccol in [('RWA', 'sub_rwa'), ('MtM', 'sub_mtm'), ('LCR (inv.)', 'sub_lcr'),
                                ('Off-Basel', 'sub_risk'), ('Disc.', 'sub_disc')]:
            check_cell(f'Table 7.2, {r["Issuer"]}, {doccol}', r.get(doccol), src[srccol],
                       'classification_scorecard.csv')
        check_cell(f'Table 7.2, {r["Issuer"]}, Score', r.get('Score (band)'), src['Score'],
                   'classification_scorecard.csv')

# ---- Table 7.3: structural events ----
if '7.3' in docx_tables:
    for r in row_dicts('7.3'):
        date, coin = r['Date'], r['Coin']
        rows = struct_events[(struct_events.Date == date) &
                              (struct_events.Coin.str.contains(coin.split(' / ')[0]))]
        if rows.empty:
            flag(f'Table 7.3, {date}', r.get('Detail'), None,
                 'structural_events_detected.csv', 'NOT_FOUND_IN_OUTPUTS')

# ---- Appendix A.1 (USDC), A.2 (USDT), A.3 (USDP) ----
usdc_assets = descriptive.set_index('Quarter')['USDC_Assets_bn']
usdc_circ = descriptive.set_index('Quarter')['USDC_Supply_bn']
usdt_assets = descriptive.set_index('Quarter')['USDT_Assets_bn']
usdt_circ = descriptive.set_index('Quarter')['USDT_Supply_bn']
lcr_ts_full = load('lcr_timeseries.csv')
l40_full = lcr_ts_full[lcr_ts_full.Outflow_pct == 40].set_index('Quarter')

if 'A.1' in docx_tables:
    for r in row_dicts('A.1'):
        q = r['Quarter']
        check_cell(f'Appendix A.1, {q}, Assets', r.get('Assets ($bn)'),
                   usdc_assets.get(q), 'descriptive_table.csv')
        cb_row = cb_ts[(cb_ts.Issuer == 'USDC') & (cb_ts.Quarter == q)]
        if not cb_row.empty:
            check_cell(f'Appendix A.1, {q}, RWA', r.get('RWA ($bn)'),
                       cb_row.iloc[0]['RWA_bn'], 'capital_bridge_timeseries.csv')

if 'A.2' in docx_tables:
    for r in row_dicts('A.2'):
        q = r['Quarter']
        check_cell(f'Appendix A.2, {q}, Circ.', r.get('Circ. ($bn)'),
                   usdt_circ.get(q), 'descriptive_table.csv')

if 'A.3' in docx_tables:
    for r in row_dicts('A.3'):
        q = r['Quarter']
        row = usdp_rwa[usdp_rwa.Quarter == q]
        if row.empty:
            continue
        row = row.iloc[0]
        check_cell(f'Appendix A.3, {q}, Circ.', r.get('Circ. ($m)'),
                   row['USDP_Circulation'] / 1e6, 'usdp_rwa_table.csv')

# ---- Appendix A.5: Bank CET1 vs RW-eq ----
if 'A.5' in docx_tables:
    bb = bank_bench.set_index('Quarter') if 'Quarter' in bank_bench.columns else None
    for r in row_dicts('A.5'):
        q = r['Quarter']
        if bb is not None and q in bb.index:
            for doccol, srccol in [('JPM CET1', 'JPM_CET1_pct'), ('BNY CET1', 'BNY_CET1_pct')]:
                if srccol in bb.columns:
                    check_cell(f'Appendix A.5, {q}, {doccol}', r.get(doccol),
                               bb.loc[q, srccol], 'bank_benchmark.csv')

print(f"Table-level comparison done: {len(discrepancies)} discrepancies so far")

# ==============================================================================
# 4. Inline-number scan: $X.XXbn, XX.X%, +/-XXXbp, bare scores
#    Best-effort — matched only against numbers_manifest.csv (curated headline
#    figures). Anything not found there is flagged NOT_FOUND_IN_OUTPUTS so a
#    human can judge whether it needs checking at all (most such numbers are
#    ordinary prose: dates, citation years, section numbers, unrelated stats).
# ==============================================================================

PATTERNS = {
    'dollar_bn': re.compile(r'\$\d[\d,]*\.?\d*\s?(?:bn|billion|m|million)'),
    'percent': re.compile(r'[−-]?\d+\.\d+\s?%'),
    'bp': re.compile(r'[−+-]?\d+\s?(?:bp|basis points)'),
    'score': re.compile(r'\bscores?\s+(?:of\s+)?[−-]?\d+\.\d+\b|\b[−-]?\d+\.\d+\s*\((?:Narrow-bank|MMF-like)\)'),
}

# 2026-07-12: discrepancy adjudication -- 'parameter' rows are design inputs
# (shock grids, WAM sweeps, the replay tuple, outflow scenarios): whitelisted
# outright, since a parameter value appearing in a caption is expected and
# correct by construction, not a measured figure to trace to a CSV cell.
param_values = set()
result_values = []
if not manifest.empty:
    for _, r in manifest.iterrows():
        if r.get('row_type') == 'parameter':
            param_values.add(round(float(r['value']), 4))
            param_values.add(round(abs(float(r['value'])), 4))
        else:
            result_values.append((r['metric'], float(r['value']), r['unit']))

inline_checked = 0
inline_not_found = 0
seen_notfound = {}  # raw value string -> list of caption labels it appeared in (for dedup)
caption_paras = [t for (kind, style, t) in items if kind == 'para' and
                  re.match(r'^Figure \d+\.\d+\.|^Figure A\.\d+\.', t)]

for cap in caption_paras:
    fig_label = re.match(r'^(Figure [\dA]+\.\d+)\.', cap).group(1)
    found_numbers = set()
    for pname, pat in PATTERNS.items():
        for m in pat.finditer(cap):
            found_numbers.add(m.group())
    for raw in found_numbers:
        v = num(raw)
        if v is None:
            continue
        inline_checked += 1
        if round(v, 4) in param_values or round(abs(v), 4) in param_values:
            continue  # whitelisted design parameter
        match = None
        for metric, mval, unit in result_values:
            tol = max(abs(mval) * 0.02, 0.05)
            if close(v, mval, tol) or close(-v, mval, tol):
                match = (metric, mval)
                break
        if match is None:
            seen_notfound.setdefault(raw, []).append(fig_label)

# dedupe: one flagged row per unique literal value, listing every caption it
# appeared in, instead of one row per (caption, value) occurrence
for raw, labels in seen_notfound.items():
    inline_not_found += 1
    locs = ', '.join(sorted(set(labels)))
    flag(f'{locs} caption(s), inline value {raw!r}', raw, None,
         'numbers_manifest.csv (no match found)', 'NOT_FOUND_IN_OUTPUTS')

print(f"Inline caption scan: {inline_checked} numbers checked, "
      f"{inline_not_found} unique unmatched values")

# ==============================================================================
# 5. Write report
# ==============================================================================
report = pd.DataFrame(discrepancies, columns=['location_in_docx', 'thesis_value',
                                               'source_value', 'source_file', 'severity'])
report.to_csv(REPORT_PATH, index=False)
print(f"\nWrote {REPORT_PATH} with {len(report)} rows")
print(report['severity'].value_counts())
