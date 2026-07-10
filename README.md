# Capital Adequacy and Liquidity Resilience of Fiat-Backed Stablecoins

**Master's Thesis — Frankfurt School of Finance & Management**  
**Author:** Luis Enrique Infante Altamirano  
**Supervisor:** Prof. Co-Pierre Georg  
**Scope:** Q3 2022 – Q4 2025 (14 quarters) · USDC (Circle) · USDT (Tether) · USDP (Paxos, narrow-bank control)

---

## What this repository is

This repository contains the complete analytical pipeline for the thesis. One Python script
(`analysis.py`) loads primary data from the master dataset, runs all four research-question
analyses end-to-end, and writes every table and figure used in the thesis. Nothing is
hard-coded or hand-tuned; all results are reproducible from the input files.

**Central finding:** asset composition and disclosure quality — not issuer size — determine
a stablecoin's regulatory profile. USDT (~$186bn) holds ~22% of its reserves in off-Basel
assets (Bitcoin, gold, secured loans), giving it an MMF-like risk character. USDC (~$75bn)
holds essentially 100% HQLA and scores as narrow-bank-like — so the larger issuer is the
riskier one, the opposite of what a size-based rule would predict. USDP is the narrow-bank
control: whole-book HQLA, near-zero rate risk, commercially wound down by end of scope.

---

## Research questions and headline results

| # | Question | Method | Headline result |
|---|---|---|---|
| **RQ1** | How does risk-weighted asset (RWA) density evolve? | Basel III risk-weight ladder (SCO60, CRE20) applied to quarterly reserve books | USDC 3.3% mean density; USDT 58.6% mean — driven by Bitcoin at 1250% (73.9% of USDT's total RWA) |
| **RQ2** | How sensitive are reserves to rate moves? | Duration-based MtM stress, 4 WAM assumptions × 3 shocks | Sovereign losses <2% of assets for all three at +400bp/180d. Short WAM is the safeguard. Rate stress ≠ total stress for USDT: BTC and gold are excluded (see below) |
| **RQ3** | Can each issuer meet a large redemption run? | Basel III LCR analogue, 20/40/60% outflow scenarios, look-through and opaque-MMF treatments | Liquidity not the binding constraint under look-through. Disclosure opacity has a measurable price: opaque USDC fails a 60% run it passes under look-through |
| **RQ4** | Narrow bank or money-market fund? | Five-axis composite scorecard (0 = narrow-bank, 100 = MMF-like), OECD/JRC-grounded weights | USDP 25 → USDC 16 → USDT 55. Ordinal ranking survives 92.5% of all possible weightings (20,000-draw Monte-Carlo) |

---

## On Bitcoin and gold — what this thesis does and does not cover

This is the most important scope note in the repository.

**What is included:**  
Bitcoin (BTC) and gold appear in the analysis at their Basel III risk weights:
- Bitcoin: **1250%** (Basel SCO60.108, Group 2 unbacked cryptoasset). At Q4 2025, Tether holds
  $8.4bn BTC → generates $105.4bn of RWA, which is 73.9% of Tether's entire RWA. This is the
  single most consequential number in RQ1.
- Gold: **100%** (commodity, no HQLA recognition, no Basel collateral treatment for a non-bank).
  Tether holds $17.5bn gold at Q4 2025.
- Secured loans: **100%** (private, unrated, no Basel collateral recognition).

These weights drive USDT's 74% RWA density and its classification as MMF-like in RQ4.
The 1250% BTC weight also creates a **procyclicality** effect: when BTC rallies, Tether's
RWA rises even without new purchases. USDT's RWA ratio peaked at ~86% (Q3 2025, near a
BTC high) and fell to 74% in Q4 2025 as BTC corrected — entirely driven by price, not
by changes in portfolio composition.

**What is NOT included — the open point:**  
RQ2 (rate stress) and RQ3 (liquidity stress) model only the **sovereign fixed-income sleeve**
(T-bills, repos). BTC and gold do not have duration — they have **price risk** — and a
duration model cannot capture that. A price-shock stress (e.g. BTC −50%/−65%/−80%;
gold −10%/−20%/−30%) against the thin reserve equity buffer would complete the risk
picture for USDT and is the **one remaining analytical extension** not yet implemented.
At Q4 2025, a −65% BTC drawdown on $8.4bn of BTC notional would produce a ~$5.5bn loss
against a ~$6.4bn reserve surplus — nearly wiping the buffer. USDC and USDP are immune
(zero BTC/gold). This gap is documented in the thesis as a limitation and a candidate
for future work.

---

## Capital adequacy — reserve level only

Capital is measured at **reserve level** (reserve assets − circulation), not group level.
A token-holder's legal claim attaches to the reserve entities, not to the parent group.
Group equity (Circle's stockholders' equity from SEC filings; Tether's consolidated group
equity including Tether Investments in BTC mining, AI, and energy) is legally walled off
from the reserve and is not counted as token-holder protection.

This is confirmed by Circle's own audited balance sheet (S-1/A, filed 2025-05-27),
which segregates "cash and cash equivalents for the benefit of stablecoin holders"
($43.9bn FY2024) separately from the parent's corporate assets. The reserve surplus
(~$191m on $43.9bn = ~4 bps at FY2024) is the actual buffer.

**Reserve-level risk-weighted equity ratios (14-quarter means):**

| Issuer | Reserve equity / RWA | What it means |
|---|---|---|
| USDC | ~6.2% mean (3–12% range) | Near-zero reserve equity; Circle's yield goes to the parent as "reserve income" — not retained in the fund |
| USDT | ~6.7% mean (1–11% range) | Tether earned $12.1bn reserve profit in 2024 but dividended ~$10.3bn to the group; reserve equity barely moved |
| USDP | ~11.5% mean (high variance) | Denominator collapse during wind-down inflates the ratio mechanically (cash migration to 20%-weighted deposits) |

**The structural finding:** stablecoins are near-zero-capital by construction. The Basel
capital lens is near-vacuous for these entities. Token-holder safety rests on asset quality
(RQ1) and short duration (RQ2) — not on an equity cushion. This is the thesis's key
capital-chapter result.

### Where the reserve yield goes — and a disclosure asymmetry

Neither issuer retains earnings in the reserve, but by different mechanisms:

- **Tether** earns profit *in* the reserve entities, then dividends it *out* to the group.
  In FY2024 the reserve earned ~$12.1bn but paid ~$10.3bn in dividends, so reserve equity
  rose only $5.2bn → $7.1bn. In FY2025 it went further — paying ~$10.9bn against ~$10.1bn
  of profit, *drawing the buffer down* to $6.3bn. The reserve risk-weighted equity ratio
  consequently declines (10.6% in 2023 → ~4.4% by end-2025): the numerator is dividended
  down while the denominator (RWA) grows.
- **Circle** earns the reserve yield at the *parent*: the Circle Reserve Fund passes its
  yield to Circle as "reserve income" (~$1.66bn in FY2024), of which Circle kept only
  ~$157m as net income (the rest to distribution partners and costs). The reserve itself
  runs ~1:1, surplus ~$191m (FY2024), ~$155m (FY2025) — structurally near-zero throughout.

**A disclosure asymmetry emerges from the data itself:** Circle's group financials are
SEC-audited and published quarterly through 2025 (S-1 + 10-Qs). Tether published a
*consolidated group* balance sheet only in 2024; from 2025, after relocating from the BVI
to El Salvador, its reports cover only the reserve-issuing entity — group figures
(including the ~$12bn Tether Investments arm) are no longer disclosed. The transparency on
the more complex entity narrowed exactly as the group grew. This is itself evidence for
the disclosure axis in RQ4.

---

## Repository structure

```
/
├── analysis.py                          # Main script — runs all 7 sections end to end
├── USDC_USDT_USDP_Basel3_Master_v8.xlsx # Master dataset (primary source of truth)
├── TB3MS.csv                            # FRED: 3-month T-bill rate, monthly 1934–present
├── MSPD_SumSecty_20010131_20260531.csv  # US Treasury MSPD: marketable bills outstanding
├── requirements.txt                     # Python dependencies
├── README.md                            # This file
└── outputs/                             # Auto-created on first run; all figures and CSVs
```

---

## Data sources

All inputs are original primary public sources. No data was purchased or licensed.

| Input | What it contains | Source |
|---|---|---|
| **USDC reserve examinations** | Monthly reserve composition (T-bills, repos, MMF, cash), circulation. Grant Thornton 2020–2022; Deloitte from Jan 2023 | [circle.com/transparency](https://www.circle.com/en/usdc#transparency) |
| **USDT consolidated reserve reports** | Quarterly ISAE 3000R opinions by BDO. Reserve assets by line (T-bills, repos, MMF, cash, BTC, gold, secured loans, other). From 2024: consolidated group balance sheet + change-in-equity | [tether.to/transparency](https://tether.to/en/transparency/) |
| **USDP reserve examinations** | Quarterly examinations by Withum (to 2024), KPMG (2025). Full asset breakdown: T-bills, repos, cash | [paxos.com/attestations](https://paxos.com/attestations/) |
| **TB3MS** | 3-month T-bill, monthly, 1934–present. Used for +400bp episode frequency analysis (90-year record) | [fred.stlouisfed.org/series/TB3MS](https://fred.stlouisfed.org/series/TB3MS) |
| **DGS1MO, DGS3MO, DGS6MO, DGS1** | Daily constant-maturity yields. In-sample yield curve for rate shock calibration | FRED, series DGS1MO / DGS3MO / DGS6MO / DGS1 |
| **MSPD** | US Treasury Monthly Statement of the Public Debt: marketable T-bills outstanding by month. Used for stablecoin T-bill footprint analysis | [fiscaldata.treasury.gov](https://fiscaldata.treasury.gov/datasets/monthly-statement-public-debt/) |
| **USDC/USDT daily price and supply** | Daily peg price and circulating supply. Used for data-driven structural-event detection | [CoinGecko](https://www.coingecko.com) |
| **Circle group financials** | Audited consolidated income statement and balance sheet 2022–2024 (S-1) and quarterly 2025 (10-Q), including reserve income, net income, stockholders' equity. Confirms reserve-level methodology and provides the group-equity context series | SEC EDGAR, CIK 0001876042 — [S-1/A filed 2025-05-27](https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001876042); 10-Q Q1/Q2/Q3 2025 |

The master Excel (`USDC_USDT_USDP_Basel3_Master_v8.xlsx`) consolidates all of the above
into 13 sheets, including `Tether_Equity_Profit` (reserve-entity and group equity/profit
from BDO consolidated attestations, 2024–2025; not disclosed pre-2024) and
`Circle_Equity_Profit` (group equity, net income and reserve income from SEC filings —
S-1 audited 2022–2024, 10-Q 2025). It is the single verified input file; no data is
fetched at runtime.

---

## How to run

### Requirements

Python 3.10 or later.

```bash
pip install -r requirements.txt
```

### Configuration

Open `analysis.py` and set `BASE_DIR` at line 65 to the folder containing the three input files:

```python
# Windows
BASE_DIR = Path(r'C:\Users\yourname\Documents\thesis')
# macOS / Linux
BASE_DIR = Path('/Users/yourname/thesis')
```

Or use an environment variable without touching the code:

```bash
# macOS / Linux
export THESIS_DIR=/path/to/folder
python analysis.py

# Windows PowerShell
$env:THESIS_DIR = "C:\path\to\folder"
python analysis.py
```

### Run

```bash
python analysis.py
```

The script prints progress to the console. All 33 outputs are written to `outputs/`
inside the base folder. A full run takes approximately 2–4 minutes.

---

## Script structure

`analysis.py` is one growing file with seven sections:

| Section | Content |
|---|---|
| 0 | Setup, imports, configuration, constants |
| 1 | Data loading and quarterly panel construction |
| 2 | Descriptive table and asset-composition figures |
| 3 | RWA computation — RQ1 (Basel III risk-weight ladder) |
| 4 | Interest-rate sensitivity — RQ2 (duration/MtM stress + 90-year rate history) |
| 5 | Redemption/LCR stress — RQ3 (look-through and opaque-MMF treatments) |
| 6 | Narrow-bank vs MMF classification — RQ4 (scorecard + weight robustness) |
| 7 | Data-driven structural-event detection (validation against known crises) |

---

## Outputs

All 31 files are written to `BASE_DIR/outputs/`.

### Figures (16)

| File | What it shows |
|---|---|
| `fig_week1_descriptive.png` | Asset composition and supply trajectories, all three issuers |
| `rwa_evolution.png` | RWA by asset class, stacked, all quarters (RQ1) |
| `rwa_ratio.png` | RWA density: look-through vs opaque-MMF sensitivity (USDC) |
| `rwa_ratio_threeway.png` | Three-issuer RWA ratio comparison over 14 quarters |
| `fig_duration_stress.png` | MtM loss heatmap: WAM × rate shock × issuer — sovereign book only (RQ2) |
| `fig_duration_timeseries.png` | Sovereign MtM loss time-series, adverse ceiling (180d/+400bp) |
| `fig_rate_history.png` | In-sample yield curve 2022–2025 (DGS series) |
| `fig_rate_history_long.png` | 90-year TB3MS history — +400bp episodes flagged and counted |
| `fig_lcr_stress.png` | LCR grid: outflow scenario × quarter × issuer (RQ3) |
| `fig_lcr_effective.png` | Effective LCR: combined RQ2×RQ3 forced-sale formula |
| `fig_lcr_vs_rate_combined.png` | RQ2×RQ3 intersection: forced-sale MtM loss vs reserve equity |
| `fig_tbill_footprint.png` | Stablecoin sector share of total US marketable T-bills outstanding |
| `fig_classification_spectrum.png` | Narrow-bank ↔ MMF positioning, all three issuers (RQ4) |
| `fig_weight_sensitivity.png` | Scorecard robustness: scheme dot-plot + 20k Monte-Carlo survival rates |
| `fig_capital_ratio.png` | Reserve-level risk-weighted equity ratio over 14 quarters (all issuers) |
| `fig_structural_events.png` | Data-driven event detection vs known crises (FTX, SVB, Paxos/BUSD) |

### Tables — CSV (15)

| File | Contents |
|---|---|
| `descriptive_table.csv` | 14-quarter supply, assets, equity buffer per issuer |
| `rwa_table.csv` | RWA by asset class, all issuers, all quarters (RQ1) |
| `usdp_rwa_table.csv` | USDP-specific RWA detail, all quarters |
| `mtm_loss_grid.csv` | MtM loss grid: WAM × shock × issuer, Q4 2025 (RQ2) |
| `mtm_loss_timeseries.csv` | MtM loss time-series under the adverse ceiling |
| `rate_shock_episodes.csv` | Historical +400bp episodes identified in TB3MS (1934–present) |
| `lcr_stress_grid.csv` | LCR grid: all scenarios, latest quarter (RQ3) |
| `lcr_timeseries.csv` | LCR time-series: all outflow scenarios, all quarters |
| `lcr_effective_coverage.csv` | Effective LCR (RQ2×RQ3 forced-sale formula) |
| `stablecoin_tbill_share.csv` | Sector T-bill share of total marketable bills outstanding |
| `classification_scorecard.csv` | RQ4: raw values, sub-scores, blended score, sensitivity band, class |
| `classification_weight_sensitivity.csv` | Scorecard under base / equal / composition-only / leave-one-out |
| `capital_bridge.csv` | Reserve-level capital adequacy test, latest quarter |
| `capital_bridge_timeseries.csv` | Reserve RW-equity ratio, all three issuers, all 14 quarters |
| `tether_capital_ratio.csv` | Tether risk-weighted equity ratio (reserve + group) and profit series, per quarter |
| `circle_equity_profit.csv` | Circle group equity, net income and reserve income (SEC filings) |
| `structural_events_detected.csv` | Data-driven structural events: date, type, coin, magnitude |

---

## Methodology notes

### Risk weights

| Asset | Risk weight | Basel reference |
|---|---|---|
| US T-bills (direct or in CRF) | 0% | CRE20.31 |
| UST-collateralised overnight repos | 0% | CRE20.31 |
| Cash at banks / external deposits | 20% | CRE20.39 |
| Opaque MMF units (USDC sensitivity) | 20% | CRE20.39 (deposit-like treatment) |
| Gold (physical) | 100% | Commodity; no HQLA recognition |
| Secured loans (Tether) | 100% | No Basel collateral recognition for non-bank |
| Other / residual USDT | 100% | Conservative; unknown composition |
| Bitcoin | **1250%** | **SCO60.108 — Group 2 unbacked cryptoasset** |

### USDC — two treatments run in parallel

- **Primary (look-through):** Circle Reserve Fund decomposed into T-bills (0%) + repos (0%) + small cash residual (20%). External cash (20%).
- **Sensitivity (opaque):** CRF NAV treated as a single opaque MMF unit at 20% (RWA) / 50% haircut (LCR). Anchored to the 2008 Reserve Primary Fund episode and March 2020 dash-for-cash. The opaque treatment is the conservative bound; the look-through is primary because Circle publishes daily fund holdings.

### LCR — outflow scenarios

Three outflow scenarios: 20%, 40%, 60% of circulation. The 40% scenario is the primary
case; 60% is the tail / stress. HQLA treatment follows the two USDC treatments above.
LCR pass threshold: 100% (Basel standard). The look-through vs opaque gap directly
measures the price of disclosure opacity.

### Classification scorecard — weight rationale

Weights are theory-based and fixed before observing results:
**Composition (RWA + off-Basel) 50% · Liquidity 20% · Rate 15% · Disclosure 15%.**
Grounded in the OECD/JRC Handbook on Constructing Composite Indicators (Nardo et al.
2005/2008) and the CAMELS supervisory rating precedent. Robustness confirmed via
20,000-draw Dirichlet-uniform Monte-Carlo: the ordinal ranking (USDT most MMF-like /
USDC most narrow-bank) survives 92.5% of all possible weighting schemes.

### Structural event detection

Events are detected algorithmically — not hand-picked. Three signal types:
- **Depeg:** |price − $1| > 0.5% on a given day
- **Supply shock:** circulating supply falls >8% over 7 days
- **Rate shock:** rolling 12-month change in 3M yield exceeds +300bp

The known crises (FTX Nov-2022, Paxos/BUSD Feb-2023, SVB Mar-2023) validate the detector
post hoc — all three are independently recovered. The March 2023 USDC depeg (−3.4%) and
simultaneous −15% supply contraction over 7 days are the largest signals in the sample.

---

## Key results summary

| Metric | USDC | USDT | USDP |
|---|---|---|---|
| RWA density — full-sample mean | 3.3% | 58.6% | 13.0% (avg, wind-down inflated) |
| BTC/gold/loans share of assets (Q4-2025) | 0% | 22.3% ($42.9bn) | 0% |
| BTC contribution to USDT RWA | — | $105.4bn / 73.9% of RWA | — |
| MtM loss, +400bp/180d — mean | 1.6% of assets | 1.5% (sovereign only) | 0.7% |
| LCR @ 40% run, look-through — mean | 206% | 192% | 96% |
| Classification score (0–100) | **16 — Narrow-bank** | **55 — MMF-like** | **25 — Narrow-bank** |
| Reserve RW-equity ratio — mean | 6.2% | 6.7% | 11.5% |
| Reserve equity at Q4-2025 | $67m (9 bps) | $6.4bn (333 bps) | ~$0 |

**The through-line:** the same driver separates the issuers on every axis — BTC, gold and
secured loans generate USDT's high RWA density, its MMF-like classification, and its
thin reserve capital despite large absolute size. Remove those three asset lines and
USDT's risk profile collapses toward USDC's. Composition, not size.

---

## Known limitations

1. **BTC/gold price-shock stress is not implemented.** The rate stress (RQ2) covers only
   the sovereign fixed-income sleeve. A price-shock scenario for the off-Basel sleeve
   (BTC −50/−65/−80%; gold −10/−20/−30%) would complete the picture and is the one
   remaining extension. At Q4 2025, a −65% BTC shock alone would nearly match Tether's
   entire reserve equity buffer.

2. **Fire-sale / price-taker assumption.** The LCR model assumes the issuer can liquidate
   assets at prevailing market prices without moving them. Holds at current sector scale
   (~1.2–2.3% of total marketable T-bills); weakens if stablecoins approach systemic size.
   (Brunnermeier-Pedersen 2009; Greenwood-Landier-Thesmar 2015.)

3. **WAM is assumed, not disclosed.** The four-point sweep (30/60/90/180 days) is a
   sensitivity band. Issuers do not publish a single consistent portfolio WAM.

4. **Parallel rate shifts only.** A yield-curve twist or steepening would differentiate
   issuers by maturity bucket.

5. **Group equity excluded by design.** Reputational considerations (a profitable parent
   may backstop its token even without a legal obligation) are not modelled.

---

## Citation

> Infante Altamirano, L. E. (2026). *Capital Adequacy and Liquidity Resilience of
> Fiat-Backed Stablecoins: A Quarterly Stress-Testing Analysis (2022–2025).*
> Master's Thesis, Frankfurt School of Finance & Management.
> Supervisor: Prof. Co-Pierre Georg.

---

## License

Code: MIT License.  
Data compiled from primary public disclosures (Circle/Tether/Paxos attestations, FRED,
US Treasury MSPD, SEC EDGAR, CoinGecko). See individual source terms for data use.
