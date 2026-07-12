"""
cross_check_thesis.py
======================
Cross-checks every numeric table cell (and a best-effort scan of inline
narrative numbers) in Infante_Thesis_Consolidated_Revision_v2.docx against
the analysis pipeline's output CSVs (outputs/) and numbers_manifest.csv.

READ-ONLY: never modifies the docx. Writes discrepancy_report.csv at the
repo root.

Usage: python cross_check_thesis.py
"""
import re
import io
import sys
import json
from pathlib import Path

import docx
import pandas as pd
from docx.document import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

REPO = Path(r"C:\Users\lucho\OneDrive\Escritorio\Thesis Master")
DOCX_PATH = REPO / "Infante_Thesis_Consolidated_Revision_v2.docx"
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
            for docval, srcval, colname, dp, tol in [
                (num(r.get(e_key)), src['Reserve_Equity_bn'], e_key, 3, 0.0006),
                (num(r.get(rwa_key)), src['RWA_bn'], rwa_key, 2, 0.006),
                (num(r.get(cov_key)), src['Coverage_x'], cov_key, 2, 0.006),
            ]:
                if docval is None:
                    continue
                if close(docval, round(srcval, dp), tol):
                    continue
                sev = 'ROUNDING' if close(docval, srcval, tol * 5) else 'MISMATCH'
                flag(f'Table 4.2, {q}, {colname}', docval, round(srcval, 3),
                     'capital_bridge_timeseries.csv', sev)
        # USDP coverage
        src_p = cb_ts[(cb_ts.Issuer == 'USDP') & (cb_ts.Quarter == q)]
        if not src_p.empty:
            docval = num(r.get('USDP cov.'))
            srcval = src_p.iloc[0]['Coverage_x']
            if docval is not None and not close(docval, round(srcval, 2), 0.006):
                sev = 'ROUNDING' if close(docval, srcval, 0.03) else 'MISMATCH'
                flag(f'Table 4.2, {q}, USDP cov.', docval, round(srcval, 3),
                     'capital_bridge_timeseries.csv', sev)

# ---- Table 4.3: retained-earnings attribution ----
if '4.3' in docx_tables:
    rmap = {(r['Issuer'], r['Year']): r for r in retained.to_dict('records')}
    for r in row_dicts('4.3'):
        key = (r['Issuer'], r['Year'])
        if key not in rmap:
            continue
        src = rmap[key]
        docval = num(r.get('Retention'))
        srcval = src['Retention_pct']
        if pd.isna(srcval) or docval is None:
            continue
        # some rows in this table are rounded to the nearest integer percent
        # rather than 1 decimal (e.g. "68%" vs sibling rows "15.5%", "-7.4%")
        if not close(docval, round(srcval, 1), 0.06) and not close(docval, round(srcval, 0), 0.06):
            flag(f'Table 4.3, {r["Issuer"]} {r["Year"]}, Retention', docval, round(srcval, 2),
                 'retained_earnings_attribution.csv', 'MISMATCH')
        elif not close(docval, round(srcval, 1), 0.06):
            flag(f'Table 4.3, {r["Issuer"]} {r["Year"]}, Retention', docval, round(srcval, 2),
                 'retained_earnings_attribution.csv', 'ROUNDING')

# ---- Table 5.1: historical rate-shock episodes ----
if '5.1' in docx_tables:
    for r, (_, src) in zip(row_dicts('5.1'), eps.iterrows()):
        docv = num(r.get('Peak rise (bp)'))
        if not close(docv, src['peak_bp'], 0.6):
            flag(f'Table 5.1, {r["Episode"]}, Peak rise', r.get('Peak rise (bp)'),
                 src['peak_bp'], 'rate_shock_episodes.csv', 'MISMATCH')
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
        docv = num(r.get('Mildest cell (30d, +100bp)'))
        if not close(docv, round(mild['MtM_Loss_pct_of_assets'], 3), 0.0006):
            flag(f'Table 5.2, {iss}, Mildest cell', r.get('Mildest cell (30d, +100bp)'),
                 mild['MtM_Loss_pct_of_assets'], 'mtm_loss_grid.csv', 'ROUNDING')
        docv2 = num(r.get('Worst cell (180d, +400bp)'))
        if not close(docv2, round(worst['MtM_Loss_pct_of_assets'], 3), 0.0006):
            flag(f'Table 5.2, {iss}, Worst cell %', r.get('Worst cell (180d, +400bp)'),
                 worst['MtM_Loss_pct_of_assets'], 'mtm_loss_grid.csv', 'ROUNDING')

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
            srcv = row.iloc[0]['LCR_pct']
            docv = num(r.get(col))
            if not close(docv, round(srcv, 1), 0.06):
                flag(f'Table 6.1, {r["Issuer"]}/{r["Treatment"]}, {col}', r.get(col),
                     round(srcv, 1), 'lcr_stress_grid.csv', 'ROUNDING')

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
            docv = num(r.get(docc))
            srcv = src[srccol]
            if not close(docv, round(srcv, 1), 0.06):
                flag(f'Table 6.2, {q}, {docc}', r.get(docc), round(srcv, 1),
                     'lcr_effective_coverage.csv', 'ROUNDING')

# ---- Table 6.3: combined-stress timeline (sleeve/buffer/break-even/replay) ----
if '6.3' in docx_tables:
    for r in row_dicts('6.3'):
        q = r['Quarter']
        if q not in sp.index or q not in cs_ts.index:
            continue
        srow, crow = sp.loc[q], cs_ts.loc[q]
        # sleeve
        docv = num(r.get('Sleeve ($bn)'))
        if docv is not None:
            srcv = round(float(srow['BTC_Notional_bn']) + float(srow['Gold_Notional_bn']), 2)
            if not close(docv, srcv, 0.006):
                flag(f'Table 6.3, {q}, Sleeve', r.get('Sleeve ($bn)'), srcv,
                     'btc_gold_shock_panel.csv', 'ROUNDING')
        # buffer
        docb = num(r.get('Buffer ($bn)'))
        srcb = round(float(srow['Reserve_Equity_bn']), 2)
        if not close(docb, srcb, 0.006):
            flag(f'Table 6.3, {q}, Buffer', r.get('Buffer ($bn)'), srcb,
                 'btc_gold_shock_panel.csv', 'ROUNDING')
        # break-even
        be = srow['Breakeven_Combined_pct']
        docbe = num(r.get('Break-even shock'))
        if pd.notna(be) and docbe is not None:
            if round(-be) != round(docbe):
                sev = 'ROUNDING' if abs(-be - docbe) < 1 else 'MISMATCH'
                flag(f'Table 6.3, {q}, Break-even shock', r.get('Break-even shock'),
                     f'{-be:.1f}%', 'btc_gold_shock_panel.csv', sev)
        # Q2-2022 replay residual
        docr = num(r.get('Q2-2022 replay residual ($bn)'))
        srcr = round(float(crow['Replay_2022Q2_Buffer_after_bn']), 2)
        if not close(docr, srcr, 0.006):
            flag(f'Table 6.3, {q}, Replay residual', r.get('Q2-2022 replay residual ($bn)'),
                 srcr, 'combined_stress_timeseries.csv', 'ROUNDING')

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
        docloss = num(r.get('Loss on Q4-2025 book ($bn)'))
        if not close(docloss, round(src['Loss_bn'], 2), 0.006):
            flag(f'Table 6.4, {r["Episode"]}, Loss', r.get('Loss on Q4-2025 book ($bn)'),
                 src['Loss_bn'], 'joint_stress_replays.csv', 'ROUNDING')
        docbuf = num(r.get('Residual buffer ($bn)'))
        if not close(docbuf, round(src['Buffer_after_bn'], 2), 0.006):
            flag(f'Table 6.4, {r["Episode"]}, Residual buffer', r.get('Residual buffer ($bn)'),
                 src['Buffer_after_bn'], 'joint_stress_replays.csv', 'ROUNDING')

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
        pairs = [(raw[1], 'raw_rwa_pct', 2, 'RWA density (mean)'),
                 (raw[2], 'raw_mtm_pct', 3, 'MtM adverse ceiling (mean)'),
                 (raw[4], 'raw_lcr_pct', 1, 'LCR at 40% run (mean)'),
                 (raw[5], 'raw_risk_pct', 2, 'Off-Basel share (mean)'),
                 (raw[6], 'raw_disc', 2, 'Disc.')]
        for docraw, srccol, dp, label in pairs:
            docv = num(docraw)
            srcv = round(src[srccol], dp)
            if docv is not None and not close(docv, srcv, 0.6 * 10 ** -dp):
                flag(f'Table 7.1, {issuer}, {label}', docraw, srcv,
                     'classification_scorecard.csv', 'ROUNDING')

# ---- Table 7.2: sub-scores + blended score ----
if '7.2' in docx_tables:
    smap = {r['Issuer']: r for r in score.to_dict('records')}
    for r in row_dicts('7.2'):
        src = smap.get(r['Issuer'])
        if src is None:
            continue
        for doccol, srccol in [('RWA', 'sub_rwa'), ('MtM', 'sub_mtm'), ('LCR (inv.)', 'sub_lcr'),
                                ('Off-Basel', 'sub_risk'), ('Disc.', 'sub_disc')]:
            docv = num(r.get(doccol))
            srcv = src[srccol]
            if docv is not None and docv != srcv:
                flag(f'Table 7.2, {r["Issuer"]}, {doccol}', r.get(doccol), srcv,
                     'classification_scorecard.csv', 'MISMATCH')
        docscore = num(r.get('Score (band)'))
        if docscore is not None and docscore != src['Score']:
            flag(f'Table 7.2, {r["Issuer"]}, Score', r.get('Score (band)'), src['Score'],
                 'classification_scorecard.csv', 'MISMATCH')

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
        docv = num(r.get('Assets ($bn)'))
        srcv = round(usdc_assets.get(q, float('nan')), 2)
        if pd.notna(srcv) and not close(docv, srcv, 0.006):
            flag(f'Appendix A.1, {q}, Assets', r.get('Assets ($bn)'), srcv,
                 'descriptive_table.csv', 'ROUNDING')
        cb_row = cb_ts[(cb_ts.Issuer == 'USDC') & (cb_ts.Quarter == q)]
        if not cb_row.empty:
            docrwa = num(r.get('RWA ($bn)'))
            srcrwa = round(cb_row.iloc[0]['RWA_bn'], 2)
            if not close(docrwa, srcrwa, 0.006):
                flag(f'Appendix A.1, {q}, RWA', r.get('RWA ($bn)'), srcrwa,
                     'capital_bridge_timeseries.csv', 'ROUNDING')

if 'A.2' in docx_tables:
    for r in row_dicts('A.2'):
        q = r['Quarter']
        docv = num(r.get('Circ. ($bn)'))
        srcv = round(usdt_circ.get(q, float('nan')), 2)
        if pd.notna(srcv) and not close(docv, srcv, 0.006):
            sev = 'ROUNDING' if close(docv, srcv, 0.02) else 'MISMATCH'
            flag(f'Appendix A.2, {q}, Circ.', r.get('Circ. ($bn)'), srcv,
                 'descriptive_table.csv', sev)

if 'A.3' in docx_tables:
    for r in row_dicts('A.3'):
        q = r['Quarter']
        row = usdp_rwa[usdp_rwa.Quarter == q]
        if row.empty:
            continue
        row = row.iloc[0]
        docv = num(r.get('Circ. ($m)'))
        srcv = round(row['USDP_Circulation'] / 1e6, 1)
        if not close(docv, srcv, 0.06):
            sev = 'ROUNDING' if close(docv, srcv, 0.5) else 'MISMATCH'
            flag(f'Appendix A.3, {q}, Circ.', r.get('Circ. ($m)'), srcv,
                 'usdp_rwa_table.csv', sev)

# ---- Appendix A.5: Bank CET1 vs RW-eq ----
if 'A.5' in docx_tables:
    bb = bank_bench.set_index('Quarter') if 'Quarter' in bank_bench.columns else None
    for r in row_dicts('A.5'):
        q = r['Quarter']
        if bb is not None and q in bb.index:
            for doccol, srccol in [('JPM CET1', 'JPM_CET1_pct'), ('BNY CET1', 'BNY_CET1_pct')]:
                docv = num(r.get(doccol))
                if srccol in bb.columns:
                    srcv = bb.loc[q, srccol]
                    if not close(docv, round(srcv, 1), 0.06):
                        flag(f'Appendix A.5, {q}, {doccol}', r.get(doccol), srcv,
                             'bank_benchmark.csv', 'ROUNDING')

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

manifest_values = []
if not manifest.empty:
    for _, r in manifest.iterrows():
        manifest_values.append((r['metric'], float(r['value']), r['unit']))

inline_checked = 0
inline_not_found = 0
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
        match = None
        for metric, mval, unit in manifest_values:
            tol = max(abs(mval) * 0.02, 0.05)
            if close(v, mval, tol) or close(-v, mval, tol):
                match = (metric, mval)
                break
        if match is None:
            inline_not_found += 1
            flag(f'{fig_label} caption, inline value {raw!r}', raw, None,
                 'numbers_manifest.csv (no match found)', 'NOT_FOUND_IN_OUTPUTS')

print(f"Inline caption scan: {inline_checked} numbers checked, {inline_not_found} unmatched")

# ==============================================================================
# 5. Write report
# ==============================================================================
report = pd.DataFrame(discrepancies, columns=['location_in_docx', 'thesis_value',
                                               'source_value', 'source_file', 'severity'])
report.to_csv(REPORT_PATH, index=False)
print(f"\nWrote {REPORT_PATH} with {len(report)} rows")
print(report['severity'].value_counts())
