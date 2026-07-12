# -*- coding: utf-8 -*-
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import pandas as pd
import importlib.util

REPO = r"C:\Users\lucho\OneDrive\Escritorio\Thesis Master"
OUT = REPO + r"\outputs"

# ---- Import analysis.py as a module (module-level defs only; __main__ guards
#      prevent main()/make_thesis_figures() from re-running on import). ----
spec = importlib.util.spec_from_file_location("analysis", REPO + r"\analysis.py")
analysis = importlib.util.module_from_spec(spec)
spec.loader.exec_module(analysis)

rows = []
def add(metric, value, unit, source_file, source_col, thesis_loc, rounding, row_type='result'):
    rows.append(dict(metric=metric, value=value, unit=unit, source_file=source_file,
                      source_column_or_cell=source_col, thesis_location=thesis_loc,
                      rounding_convention=rounding, row_type=row_type))

def add_param(metric, value, unit, const_name, thesis_loc):
    """2026-07-12: discrepancy adjudication -- design parameters (scenario
    inputs fixed before the analysis ran), not measured/computed results.
    Distinguished via row_type so the checker can whitelist them: a
    parameter appearing in a caption (e.g. '+400 bp') is expected and
    correct by construction, not something to trace to a CSV cell."""
    add(metric, value, unit, 'analysis.py', const_name, thesis_loc,
        'exact — fixed design parameter, not rounded from a computation',
        row_type='parameter')

# ------------------------------------------------------------------
# 1. Classification scores (Table 7.2 / Ch.7) + Monte Carlo (Fig 7.1 / Ch.7)
# ------------------------------------------------------------------
score = pd.read_csv(f"{OUT}/classification_scorecard.csv")
for _, r in score.iterrows():
    add(f'Classification score ({r["Issuer"]})', r['Score'], 'score (0-100)',
        'classification_scorecard.csv', 'Score', 'Chapter 7, Table 7.2',
        'rounded to 1 decimal in source')

schemes_df, mc_summary = analysis.compute_weight_sensitivity(score)
add('Monte Carlo: full ordinal claim intact', mc_summary['ordinal_intact_pct'], '%',
    'analysis.py (live computation — not persisted to any CSV)',
    "compute_weight_sensitivity() -> mc_summary['ordinal_intact_pct'], n=20000, seed=0",
    'Chapter 7, Figure 7.1 / narrative', 'rounded to 1 decimal in source')
for key, label in [('usdt_highest_pct', 'Monte Carlo: USDT scores highest'),
                    ('usdc_narrowbank_pct', 'Monte Carlo: USDC stays narrow-bank (<33)'),
                    ('usdp_narrowbank_pct', 'Monte Carlo: USDP stays narrow-bank (<33)'),
                    ('usdt_mmf_pct', 'Monte Carlo: USDT stays MMF-like (>=50)')]:
    add(label, mc_summary[key], '%',
        'analysis.py (live computation — not persisted to any CSV)',
        f"compute_weight_sensitivity() -> mc_summary['{key}'], n=20000, seed=0",
        'Chapter 7, Figure 7.1 / narrative', 'rounded to 1 decimal in source')

# ------------------------------------------------------------------
# 2. RW-equity: Q4-2025 snapshot + 14q means (Chapter 4, Table 4.3 / 4.5)
# ------------------------------------------------------------------
cb = pd.read_csv(f"{OUT}/capital_bridge.csv")
for _, r in cb.iterrows():
    add(f'RW-equity ratio, Q4-2025 ({r["Issuer"]})', r['RW_Equity_Ratio_pct'], '%',
        'capital_bridge.csv', 'RW_Equity_Ratio_pct', 'Chapter 4, Table 4.3 (Q4-2025 row)',
        'rounded to 2 decimals in source')

cbts = pd.read_csv(f"{OUT}/capital_bridge_timeseries.csv")
for iss, grp in cbts.groupby('Issuer'):
    mean_val = round(grp['RW_Equity_Ratio_pct'].mean(), 1)
    add(f'RW-equity ratio, 14-quarter mean ({iss})', mean_val, '%',
        'capital_bridge_timeseries.csv', 'mean(RW_Equity_Ratio_pct) over 14 quarters',
        'Chapter 4, Table 4.5 / narrative (%.1f mean)',
        'computed as simple mean of the 14 quarterly values, each already rounded to 2dp in source')

# ------------------------------------------------------------------
# 3. LCR means (Ch.6, Table 6.2/Fig 6.3) + Q4-2022 artifact
# ------------------------------------------------------------------
lcr_ts = pd.read_csv(f"{OUT}/lcr_timeseries.csv")
l40 = lcr_ts[lcr_ts['Outflow_pct'] == 40]
for col, label in [('USDC_LCR_LT', 'LCR @ 40% run, look-through, 14q mean (USDC)'),
                    ('USDT_LCR_LT', 'LCR @ 40% run, look-through, 14q mean (USDT)'),
                    ('USDP_LCR', 'LCR @ 40% run, 14q mean (USDP)')]:
    add(label, round(l40[col].mean(), 0), '%', 'lcr_timeseries.csv',
        f'mean({col}) at Outflow_pct=40 over 14 quarters', 'Chapter 6, Table 6.3 / narrative',
        'computed as simple mean; source values carry 3 decimals')

q422 = l40[l40['Quarter'] == '2022-Q4'].iloc[0]
add('LCR @ 40% run, Q4-2022 fund-transition artifact (USDC)', round(q422['USDC_LCR_LT'], 1), '%',
    'lcr_timeseries.csv', 'USDC_LCR_LT at Quarter=2022-Q4, Outflow_pct=40',
    'Chapter 6, Figure 6.3 annotation / Table 6.2', 'rounded to 1 decimal from 3dp source value')

# ------------------------------------------------------------------
# 4. MtM grid corner cells (Ch.5, Table 5.2)
# ------------------------------------------------------------------
mtm = pd.read_csv(f"{OUT}/mtm_loss_grid.csv")
for iss in ('USDC', 'USDT', 'USDP'):
    mild = mtm[(mtm.Issuer == iss) & (mtm.WAM_days == 30) & (mtm.Shock_bp == 100)].iloc[0]
    worst = mtm[(mtm.Issuer == iss) & (mtm.WAM_days == 180) & (mtm.Shock_bp == 400)].iloc[0]
    add(f'MtM loss, mildest cell 30d/+100bp ({iss})', round(mild['MtM_Loss_pct_of_assets'], 3),
        '% of assets', 'mtm_loss_grid.csv', 'MtM_Loss_pct_of_assets @ WAM_days=30, Shock_bp=100',
        'Chapter 5, Table 5.2', 'rounded to 3 decimals in source')
    add(f'MtM loss, worst cell 180d/+400bp ({iss})', round(worst['MtM_Loss_pct_of_assets'], 3),
        '% of assets', 'mtm_loss_grid.csv', 'MtM_Loss_pct_of_assets @ WAM_days=180, Shock_bp=400',
        'Chapter 5, Table 5.2', 'rounded to 3 decimals in source')

# ------------------------------------------------------------------
# 5. Break-even arc (Ch.6, Table 6.3/Fig 6.4)
# ------------------------------------------------------------------
sp = pd.read_csv(f"{OUT}/btc_gold_shock_panel.csv").drop_duplicates('Quarter').set_index('Quarter')
for q in ['2023-Q1', '2023-Q4', '2024-Q4', '2025-Q1', '2025-Q2', '2025-Q3', '2025-Q4']:
    v = sp.loc[q, 'Breakeven_Combined_pct']
    add(f'Break-even combined BTC+gold shock ({q})', round(-float(v), 0), '%',
        'btc_gold_shock_panel.csv', 'Breakeven_Combined_pct (sign-flipped to match thesis convention)',
        'Chapter 6, Table 6.3 / Figure 6.4', 'rounded to nearest integer percent in thesis text')

# ------------------------------------------------------------------
# 6. Q2-2022 replay residuals (Ch.6, Table 6.3)
# ------------------------------------------------------------------
cs = pd.read_csv(f"{OUT}/combined_stress_timeseries.csv").drop_duplicates('Quarter').set_index('Quarter')
for q in ['2023-Q4', '2025-Q4']:
    v = cs.loc[q, 'Replay_2022Q2_Buffer_after_bn']
    add(f'Q2-2022 replay residual buffer ({q})', round(float(v), 2), '$bn',
        'combined_stress_timeseries.csv', 'Replay_2022Q2_Buffer_after_bn',
        'Chapter 6, Table 6.3 narrative', 'rounded to 2 decimals in source')

# ------------------------------------------------------------------
# 7. Sleeve values (Ch.6, Table 6.3)
# ------------------------------------------------------------------
for q in ['2023-Q1', '2025-Q4']:
    r = sp.loc[q]
    sleeve = round(float(r['BTC_Notional_bn']) + float(r['Gold_Notional_bn']), 2)
    add(f'BTC+gold sleeve ({q})', sleeve, '$bn', 'btc_gold_shock_panel.csv',
        'BTC_Notional_bn + Gold_Notional_bn', 'Chapter 6, Table 6.3',
        'sum of two source columns, each rounded to 3 decimals; re-rounded to 2dp')

# ------------------------------------------------------------------
# 8. Celsius / loan-book bound (Ch.6, Section 6.3.5) — derived, two sources
# ------------------------------------------------------------------
rwa = pd.read_csv(f"{OUT}/rwa_table.csv")
loan_book = rwa[rwa.Quarter == '2025-Q4']['USDT_RWA_SecLoans'].iloc[0] / 1e9
replay_buf = cs.loc['2025-Q4', 'Replay_2022Q2_Buffer_after_bn']
celsius_bound = replay_buf / loan_book * 100
add('Celsius / loan-book impairment bound', round(celsius_bound, 2), '%',
    'rwa_table.csv + combined_stress_timeseries.csv',
    'Replay_2022Q2_Buffer_after_bn[2025-Q4] / (USDT_RWA_SecLoans[2025-Q4]/1e9) * 100',
    'Chapter 6, Section 6.3.5', 'derived ratio, not a single source cell; thesis rounds to 2 decimals (0.06%)')

# ------------------------------------------------------------------
# 9. T-bill footprint range (Ch.6, Table 6.1/Fig 6.1)
# ------------------------------------------------------------------
share = pd.read_csv(f"{OUT}/stablecoin_tbill_share.csv")
add('T-bill footprint, minimum share', round(share['Share_pct'].min(), 1), '%',
    'stablecoin_tbill_share.csv', 'min(Share_pct)', 'Chapter 6, Section 6.2 narrative / Figure 6.1',
    'rounded to 1 decimal')
add('T-bill footprint, maximum share', round(share['Share_pct'].max(), 1), '%',
    'stablecoin_tbill_share.csv', 'max(Share_pct)', 'Chapter 6, Section 6.2 narrative / Figure 6.1',
    'rounded to 1 decimal')

# ------------------------------------------------------------------
# 2026-07-12: discrepancy adjudication -- USDP panel-start/peak/panel-end
# assets (Task 2c: Figure 4.2 caption and Section 7.3.3 wind-down sentence).
# ------------------------------------------------------------------
usdp_rwa_m = pd.read_csv(f"{OUT}/usdp_rwa_table.csv")
_q3_22 = usdp_rwa_m[usdp_rwa_m.Quarter == '2022-Q3']['USDP_Total_Assets'].iloc[0]
_maxrow = usdp_rwa_m.loc[usdp_rwa_m['USDP_Total_Assets'].idxmax()]
_q4_25 = usdp_rwa_m[usdp_rwa_m.Quarter == '2025-Q4']['USDP_Total_Assets'].iloc[0]
add('USDP assets, panel start (2022-Q3)', round(_q3_22 / 1e6, 1), '$m',
    'usdp_rwa_table.csv', 'USDP_Total_Assets[2022-Q3]', 'Figure 4.2 caption; Section 7.3.3',
    'rounded to nearest $m')
add(f'USDP assets, panel maximum ({_maxrow["Quarter"]})', round(_maxrow['USDP_Total_Assets'] / 1e6, 1), '$m',
    'usdp_rwa_table.csv', 'max(USDP_Total_Assets)', 'Figure 4.2 caption (candidate reference point)',
    'rounded to nearest $m')
add('USDP assets, panel end (2025-Q4)', round(_q4_25 / 1e6, 1), '$m',
    'usdp_rwa_table.csv', 'USDP_Total_Assets[2025-Q4]', 'Figure 4.2 caption; Section 7.3.3',
    'rounded to nearest $m')

# ------------------------------------------------------------------
# 10. 2026-07-12: discrepancy adjudication -- design-parameter rows.
#     Scenario inputs fixed before the analysis ran (shock grids, WAM
#     sweeps, the historical replay tuple, LCR outflow scenarios). These
#     appear verbatim in figure captions and table headers; they are not
#     "measured" and have no CSV cell of their own to trace to.
# ------------------------------------------------------------------
for bp in analysis.SHOCKS_BP:
    add_param(f'Rate shock leg (+{bp}bp)', bp, 'bp', 'SHOCKS_BP',
              'Chapter 5, Table 5.2 grid / Figure 5.2')
for pct in analysis.PRICE_SHOCKS:
    add_param(f'Combined price shock (-{int(pct*100)}%)', -int(pct * 100), '%', 'PRICE_SHOCKS',
              'Chapter 6, Table 6.3 grid / Appendix Figure A.1')
_replay_btc, _replay_gold, _replay_bp = -57.3, -6.9, 105
add_param('Q2-2022 replay: BTC leg', _replay_btc, '%',
          "compute_combined_stress_timeseries() default arg replay=(0.573, 0.069, 105)",
          'Chapter 6, Table 6.3/6.4 narrative')
add_param('Q2-2022 replay: gold leg', _replay_gold, '%',
          "compute_combined_stress_timeseries() default arg replay=(0.573, 0.069, 105)",
          'Chapter 6, Table 6.3/6.4 narrative')
add_param('Q2-2022 replay: rate leg', _replay_bp, 'bp',
          "compute_combined_stress_timeseries() default arg replay=(0.573, 0.069, 105)",
          'Chapter 6, Table 6.3/6.4 narrative')
for d in analysis.WAM_DAYS:
    add_param(f'WAM sweep ({d}d)', d, 'days', 'WAM_DAYS',
              'Chapter 5, Table 5.2 / Figure 5.2')
for frac in analysis.OUTFLOW_SCENARIOS:
    add_param(f'Redemption outflow scenario ({int(frac*100)}%)', int(frac * 100), '%',
              'OUTFLOW_SCENARIOS', 'Chapter 6, Table 6.1/6.2, Figure 6.1')

# ------------------------------------------------------------------
# Write manifest
# ------------------------------------------------------------------
manifest = pd.DataFrame(rows, columns=['metric', 'value', 'unit', 'source_file',
                                        'source_column_or_cell', 'thesis_location',
                                        'rounding_convention', 'row_type'])
out_path = REPO + r"\numbers_manifest.csv"
manifest.to_csv(out_path, index=False)
print(f"Wrote {out_path} with {len(manifest)} rows")
print(manifest.to_string())
