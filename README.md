# Empirical Estimation of Electricity Demand of Electric Motor-Driven Systems in Industrial Unit Operations


Code and data supporting the manuscript "Closing a Motor-Power Data Gap in Early-Stage TEA and LCA Using Empirical Equipment Scaling."

This repository contains the vendor-sourced equipment dataset, the per-equipment analysis notebooks, and the standalone figure-generation script used to derive and validate a three-level empirical framework for estimating the electricity demand of electric motor-driven systems (EMDS) in five industrial unit operations: agitated tanks, mixers, filtration units, dryers, and rotary kilns.

## Repository contents

| File | Description |
|---|---|
| `Equipment Capacity Power Type Input File.xlsx` | Vendor-sourced dataset (~800 entries) with one sheet per equipment type: rated capacity, installed motor power, subtype, and (for agitated tanks) maximum shaft rotational speed |
| `Agitated_Tanks.ipynb` | Full analysis for agitated tanks (volume basis, kW·m⁻³) |
| `Mixers.ipynb` | Full analysis for mixers (volume basis, kW·m⁻³) |
| `Filtration_Units.ipynb` | Full analysis for filtration units (area basis, kW·m⁻²) |
| `Dryers.ipynb` | Full analysis for dryers (area basis, kW·m⁻²) |
| `Rotary_Kilns.ipynb` | Full analysis for rotary kilns (volume basis, kW·m⁻³); size-independent constant model |
| `combined_figures_CORRECTED.py` | Standalone script that regenerates all manuscript and SI figures at 300 DPI (PNG/PDF/TIFF) |
| `requirements.txt` | Python dependencies |

## Framework overview

The framework provides power-intensity estimates (P/Vₘ or P/A) at three levels of user information, converted to operating draw with a load factor LF = 0.60:

- **Level 1** — binned geometric means (Small/Medium/Large capacity classes) with geometric standard deviations.
- **Level 2** — log-log OLS power-law correlation with prescribed default geometry parameters.
- **Level 3** — the Level 2 correlation evaluated with user-measured geometry inputs.

Energy per batch volume follows E/Vb = LF · (P/Vₘ) · τ.

Rotary kilns are the exception: installed drum-drive power intensity is size-independent over the compiled range, so Levels 1 and 2 collapse to a single constant (geometric mean in log space) rather than a power law.

## Statistical workflow (each notebook)

1. Distributional diagnostics (Kolmogorov–Smirnov fits, Freedman–Diaconis histograms)
2. Model selection across polynomial degrees 1–3 with R², adjusted R², RMSE, AIC, BIC, nested F-tests, and a coefficient-CI eligibility criterion (all 95% CIs must exclude zero)
3. Influential-observation screening — leverage, externally studentized residuals, DFFITS, and Cook's distance for the regression-based equipment; a 1.5 × IQR rule on log₁₀(P/V) for rotary kilns
4. Refit after removal; HC3 (heteroscedasticity-robust) inference on slopes
5. Assumption checks — Shapiro–Wilk / Kolmogorov–Smirnov normality and Breusch–Pagan homoscedasticity
6. Level-1 grouping — capacity-class construction with geometric means, GSDs, 95% CIs, and smearing factors (SF = exp((ln 10)² s² / 2)) for arithmetic-mean use cases
7. Subtype analysis (where subtypes exist) — HC3 Wald tests of each subtype against the pooled model and pairwise slope contrasts, Holm-corrected

Each output-producing cell begins with a `print("generating Table Sx")` / `print("generating Figure Sx")` descriptor identifying the manuscript or Supporting Information item it produces.

## Reproducing the results

```bash
pip install -r requirements.txt
```

Run the notebooks top to bottom with the workbook in the same directory. Notebooks render all figures inline; file output (PNG/PDF/TIFF at 300 DPI) is handled exclusively by the standalone script:

```bash
python combined_figures.py
```

Figures are written to a `Figures/` subdirectory.

## Requirements

Python ≥ 3.10 with `numpy`, `pandas`, `matplotlib`, `scipy`, `statsmodels`, `openpyxl`. Exact versions are pinned in `requirements.txt`.

## Data notes

- All capacities and powers are vendor nameplate values. Literature-reported power intensities correspond to operating draw; multiply nameplate P/Vₘ by LF for comparison.
- Volume-based equipment (tanks, mixers, kilns) is reported in kW·m⁻³ against m³; area-based equipment (filtration units, dryers) in kW·m⁻² against m².

## Citation

If you use this code or dataset, please cite:

> Mostofifar, A.; Khatami, H.; Zhang, T.; Ye, T.; Jiang, D.; Closing a Motor-Power Data Gap in Early-Stage TEA and LCA Using Empirical Equipment Scaling. [Journal], [year]. DOI: [article DOI]

Archived release: Zenodo DOI: 10.5281/zenodo.XXXXXXX

## License

Code is released under the MIT License; the dataset is released under CC BY 4.0. See `LICENSE` for details.
