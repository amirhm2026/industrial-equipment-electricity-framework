"""
=============================================================================
COMBINED PUBLICATION FIGURES
=============================================================================
Standalone script that reads the equipment dataset from Excel, redraws the
required panels in the approved publication style, and exports each combined
figure at 300 DPI (PNG + PDF + TIFF).

Each figure is built in its own clearly-labeled section below. Run the whole
file to regenerate all figures, or comment out calls in MAIN you don't need.

Data source : Equipment_Capacity_Power_Type_Input_File.xlsx
Output dir  : OUTDIR below
=============================================================================
"""
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# ============================================================================= #
# SHARED SETUP  —  style, palette, data loading, helpers
# ============================================================================= #

# ---- Publication style -------------------------------------------------------
plt.style.use("default")
plt.rcParams.update({
    "font.family": "serif",
    "font.serif": ["Times New Roman", "Liberation Serif", "DejaVu Serif"],
    "font.size": 11,
    "axes.facecolor": "white",
    "axes.edgecolor": "black",
    "axes.linewidth": 0.8,
    "axes.labelcolor": "black",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.grid": False,
    "xtick.color": "black",
    "ytick.color": "black",
    "xtick.labelsize": 9,
    "ytick.labelsize": 9,
    "figure.facecolor": "white",
    "savefig.dpi": 300,
    "savefig.bbox": "tight",
    "savefig.facecolor": "white",
    "mathtext.fontset": "stix",
})

# ---- Shared colors -----------------------------------------------------------
OBS_COLOR        = "#1B3A5C"   # deep navy — observations / bars
FIT_LINE_COLOR   = "#B03A2E"   # deep red — fit lines
RMSE_BAND_COLOR  = "#64A082"   # sage green — RMSE band
GM_MARKER_COLOR  = "#B07D2A"   # amber — geometric mean markers
REMOVED_PT_COLOR = "#A0293D"   # muted red — removed influential points
REF_LINE_COLOR   = "#B07D2A"   # amber — reference lines

CAP_FACE  = {"Small": "#1B3A5C", "Medium": "#3D6B45", "Large": "#7A1A2E"}
CAP_ALPHA = {"Small": 0.28, "Medium": 0.28, "Large": 0.20}

SUBTYPE_STYLES = {
    # ── Mixers (6 subtypes) — colours span the full hue wheel ──────────────────
    "3D Drum Mixers":          {"color": "#2166AC", "marker": "o", "linestyle": "-"},      # strong blue   (~210°)
    "Nauta Mixers":            {"color": "#D73027", "marker": "s", "linestyle": "--"},     # vivid red     (~0°)
    "Horizontal Ribbon Mixers":{"color": "#1A9641", "marker": "^", "linestyle": ":"},      # forest green  (~130°)
    "V-Blender Mixers":        {"color": "#F28C00", "marker": "D", "linestyle": "-."},     # vivid orange  (~37°)
    "Double Cone Mixers":      {"color": "#7B2D8B", "marker": "*", "linestyle": (0,(4,2,1,2))},  # violet  (~285°)
    "Vertical Ribbon Mixers":  {"color": "#00A0A0", "marker": "P", "linestyle": (0,(1,1))},       # teal    (~180°)
    # ── Filtration (3 subtypes) ────────────────────────────────────────────────
    "Nutsche Filters":         {"color": "#2166AC", "marker": "o", "linestyle": "-"},      # blue
    "Rotary Drum Filters":     {"color": "#D73027", "marker": "^", "linestyle": "--"},     # red
    "Plate & Frame Filters":   {"color": "#1A9641", "marker": "s", "linestyle": ":"},      # green
    # ── Dryers (5 subtypes) — same vivid set, none close to each other ─────────
    "Tray Dryers":             {"color": "#2166AC", "marker": "o", "linestyle": "-"},      # blue     (~210°)
    "Belt Dryers":             {"color": "#1A9641", "marker": "s", "linestyle": "--"},     # green    (~130°)
    "Drum Dryers":             {"color": "#D73027", "marker": "^", "linestyle": ":"},      # red      (~0°)
    "Rotary Drum Dryers":      {"color": "#F28C00", "marker": "D", "linestyle": "-."},     # orange   (~37°)
    "Rotary Plate Dryers":     {"color": "#7B2D8B", "marker": "*", "linestyle": (0,(4,2,1,2))},   # violet (~285°)
    # ── Single-subtype equipment ───────────────────────────────────────────────
    "Agitated Tanks":          {"color": "#2166AC", "marker": "o", "linestyle": "-"},
    "Rotary Kiln":             {"color": "#2166AC", "marker": "o", "linestyle": "-"},
    "Rotary Kilns":            {"color": "#2166AC", "marker": "o", "linestyle": "-"},
}

# ---- Paths -------------------------------------------------------------------
# By default the script looks for the Excel file in the same folder as itself,
# and saves figures into a "Figures" subfolder next to it (auto-created).
# To override, set XL and OUTDIR to explicit paths below.
import os
HERE = os.path.dirname(os.path.abspath(__file__))
XL = os.path.join(HERE, "Equipment Capacity Power Type Input File.xlsx")
OUTDIR = os.path.join(HERE, "Figures")
os.makedirs(OUTDIR, exist_ok=True)
print(f"Excel input: {XL}")
print(f"Output dir : {OUTDIR}")

# ---- Unit conversions (match notebooks exactly) ------------------------------
GAL_TO_M3 = 3.78541 / 1000
HP_TO_KW  = 745.7 / 1000

def load_equipment(sheet, cap_col, pow_col, cap_in_gal, pow_in_hp):
    df = pd.read_excel(XL, sheet_name=sheet)
    # Workbook columns were renamed "... Type" -> "... Subtype"; accept either.
    type_cols = [c for c in df.columns if "Subtype" in str(c)] or \
                [c for c in df.columns if "Type" in str(c)]
    type_col = type_cols[0]
    cap  = df[cap_col].astype(float).values
    powr = df[pow_col].astype(float).values
    if cap_in_gal:
        cap = cap * GAL_TO_M3
    if pow_in_hp:
        powr = powr * HP_TO_KW
    out = pd.DataFrame({
        "Capacity": cap,
        "Power_kW": powr,
        "Type": df[type_col].astype(str).values,
    })
    out = out[out["Capacity"] > 0].copy()
    out["PV"] = out["Power_kW"] / out["Capacity"]
    out = out[out["PV"] > 0].copy()
    out["logCap"] = np.log10(out["Capacity"])
    out["logPV"]  = np.log10(out["PV"])
    out = out.sort_values("Capacity").reset_index(drop=True)
    return out

EQUIPMENT = {
    "Agitated Tanks":   (load_equipment("Agitated Tanks",   "Tank Rated Volume, gal",  "Motor Power, hp", True,  True),  "V", "-3"),
    "Mixers":           (load_equipment("Mixers",           "Mixer Rated Volume, gal", "Total Power, hp", True,  True),  "V", "-3"),
    "Filtration Units": (load_equipment("Filtration Units", "Filter Surace Area, m2",  "Total Power, hp", False, True),  "A", "-2"),
    "Dryers":           (load_equipment("Dryers",           "Heat Transfer Area, m2",  "Total Power, hp", False, True),  "A", "-2"),
    "Calciners":        (load_equipment("Rotary Kilns",     "Volumetric Capacity, m3", "Motor Power, kW", False, False), "V", "-3"),
}

# Reader-facing display names — used only where the equipment name is shown
# directly on a figure (titles, row labels). Internal EQUIPMENT dict keys,
# the Excel sheet name, SKIP_GROUPS, and outlier-method comparisons all stay
# as "Calciners" since that's the literal Excel sheet/column this data comes
# from; only the on-figure text changes to match the manuscript's naming.
DISPLAY_NAME = {"Calciners": "Rotary Kilns"}

# ---- Helpers -----------------------------------------------------------------
def fd_bins(data):
    data = np.asarray(data)
    q25, q75 = np.percentile(data, [25, 75])
    iqr = q75 - q25
    n = data.size
    if iqr == 0 or n < 2:
        return 10
    bw = 2 * iqr / (n ** (1/3))
    if bw == 0:
        return 10
    return max(5, int(np.ceil((data.max() - data.min()) / bw)))

def unit_label(sup):
    return f"kW\u00b7m$^{{{sup}}}$"

def save_figure(fig, basename):
    fig.savefig(f"{OUTDIR}/{basename}.png",  dpi=300)
    fig.savefig(f"{OUTDIR}/{basename}.pdf")
    fig.savefig(f"{OUTDIR}/{basename}.tiff", dpi=300, pil_kwargs={"compression": "tiff_lzw"})
    print(f"  saved: {basename}.png / .pdf / .tiff")


# ============================================================================= #
# FIGURE S1  —  POWER INTENSITY DISTRIBUTIONS (raw + log), 5 rows x 2 cols
#   Row = equipment type (a-e); Left col = raw histogram; Right = log10
# ============================================================================= #
def figure_s1_distributions():
    print("Figure S1 — power intensity distributions")
    order = ["Agitated Tanks", "Mixers", "Filtration Units", "Dryers", "Calciners"]
    letters = ["a", "b", "c", "d", "e"]

    fig, axes = plt.subplots(len(order), 2, figsize=(8.5, 12.5))

    for r, name in enumerate(order):
        df, kind, sup = EQUIPMENT[name]
        pv = df["PV"].values
        logpv = df["logPV"].values
        unit = unit_label(sup)

        ax = axes[r, 0]
        ax.hist(pv, bins=fd_bins(pv), color=OBS_COLOR, alpha=0.55,
                edgecolor=OBS_COLOR, linewidth=0.4)
        ax.set_xlabel(f"Power intensity [{unit}]", style="italic", fontsize=9)
        ax.set_ylabel("Count", style="italic", fontsize=9)
        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)

        ax2 = axes[r, 1]
        ax2.hist(logpv, bins=fd_bins(logpv), color=OBS_COLOR, alpha=0.55,
                 edgecolor=OBS_COLOR, linewidth=0.4)
        ax2.set_xlabel(f"log$_{{10}}$(Power intensity)", style="italic", fontsize=9)
        ax2.set_ylabel("Count", style="italic", fontsize=9)
        ax2.spines["top"].set_visible(False); ax2.spines["right"].set_visible(False)

        axes[r, 0].annotate(
            f"({letters[r]}) {DISPLAY_NAME.get(name, name)}",
            xy=(-0.28, 0.5), xycoords="axes fraction",
            ha="center", va="center", rotation=90,
            fontsize=11, color="#333333", style="italic",
        )

    axes[0, 0].set_title("Raw", fontsize=11, color="#555555", style="italic", pad=8)
    axes[0, 1].set_title("Log-transformed", fontsize=11, color="#555555", style="italic", pad=8)

    plt.tight_layout(w_pad=2.0, h_pad=1.8)
    plt.subplots_adjust(left=0.13)
    save_figure(fig, "Figure_S1_distributions")
    plt.close(fig)


# ============================================================================= #
# FIGURE S2  —  POWER-CAPACITY TRENDS ACROSS DEVICES, 2 rows x 3 cols
#   One panel per equipment type (raw/power scale only). Subtypes shown with
#   approved color + marker. Scatter only (no fit lines). Panels labeled a-e.
#   6th grid slot (row 2, col 3) intentionally left blank.
# ============================================================================= #
def figure_s2_power_capacity_trends():
    print("Figure S2 — power-capacity trends across devices")
    from matplotlib.lines import Line2D

    order = ["Agitated Tanks", "Mixers", "Filtration Units", "Dryers", "Calciners"]
    letters = ["a", "b", "c", "d", "e"]
    positions = [(0, 0), (0, 1), (0, 2), (1, 0), (1, 1)]  # (row, col) for each panel

    fig, axes = plt.subplots(2, 3, figsize=(15.5, 9.0))

    for (r, c), name, letter in zip(positions, order, letters):
        df, kind, sup = EQUIPMENT[name]
        unit = unit_label(sup)
        cap_word = "Volume" if kind == "V" else "Area"
        subtypes = list(df["Type"].unique())

        ax = axes[r, c]
        legend_handles = []

        for t in subtypes:
            st = SUBTYPE_STYLES.get(t, {"color": "#333333", "marker": "o"})
            sub = df[df["Type"] == t]

            ax.scatter(sub["Capacity"], sub["PV"],
                       color=st["color"], marker=st["marker"],
                       s=20, alpha=0.65, edgecolors=st["color"], linewidths=0.4, zorder=2)

            legend_handles.append(
                Line2D([0], [0], marker=st["marker"], color="w",
                       markerfacecolor=st["color"], markeredgecolor=st["color"],
                       markersize=7, label=t.replace(" Mixers", "").replace(" Filters", "")
                                             .replace(" Dryers", "").replace(" Tanks", "")
                                             .replace(" Kiln", "").replace(" Kilns", ""))
            )

        # Axes formatting
        ax.set_xscale("log"); ax.set_yscale("log")
        ax.set_xlabel(f"{cap_word} (m$^{{{'3' if kind=='V' else '2'}}}$)",
                      style="italic", fontsize=13)
        ax.set_ylabel(f"Power intensity [{unit}]", style="italic", fontsize=13)
        ax.tick_params(axis="both", labelsize=11)
        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)

        # Panel title: letter + equipment name, above the axes
        ax.set_title(f"({letter}) {DISPLAY_NAME.get(name, name)}", fontsize=14, color="#333333",
                     style="italic", pad=10, loc="left")

        # Legend — only when >1 subtype. Placed inside the panel (upper right)
        # since there is no dedicated side column to hold it anymore.
        if len(subtypes) > 1:
            leg = ax.legend(handles=legend_handles, fontsize=9.5, frameon=True,
                            facecolor="white", edgecolor="#cccccc",
                            loc="best", handletextpad=0.4, borderpad=0.5,
                            labelspacing=0.35)
            leg.set_zorder(5)

    # Leave the 6th grid slot (row 1, col 2) blank
    axes[1, 2].axis("off")

    plt.tight_layout(w_pad=2.5, h_pad=3.0)
    save_figure(fig, "Figure_S2_power_capacity_trends")
    plt.close(fig)


# ============================================================================= #
# FIGURE: INFLUENCE DIAGNOSTICS  —  one 2x2 panel figure PER SUBTYPE
#   Panels: (a) leverage, (b) studentized residual, (c) DFFITS, (d) Cook's D
#   Normal points navy, influential points red diamonds, dashed thresholds.
#   For single-type equipment (Tanks, Calciners) the one figure covers the group.
#   Each subtype is saved as its own file: Figure_influence_<equipment>_<subtype>.*
# ============================================================================= #
def _compute_influence(sub):
    """Fit degree-1 log-log OLS on a subtype subset; return diagnostics + thresholds."""
    import statsmodels.api as sm
    x = sub["logCap"].values
    y = sub["logPV"].values
    n = len(y)
    k = 1
    X = np.vander(x, k + 1, increasing=True)
    model = sm.OLS(y, X).fit()
    inf = model.get_influence()
    lev = inf.hat_matrix_diag
    sr  = inf.resid_studentized_external
    dff = inf.dffits[0]
    cks = inf.cooks_distance[0]
    mean_lev = (k + 1) / n
    thresholds = {
        "lev": 3 * mean_lev,
        "sr": 3,
        "dff": 2 * np.sqrt((k + 2) / (n - k - 2)) if n - k - 2 > 0 else np.inf,
        "cks": 4 / n,
    }
    return lev, sr, dff, cks, thresholds


def _diag_panel(ax, obs, vals, normal_mask, thresh, ylabel, thresh_label,
                zero_line=False):
    """Per-panel scatter + threshold lines + small mini-legend identifying the threshold."""
    from matplotlib.lines import Line2D
    # Normal observations
    ax.scatter(obs[normal_mask], vals[normal_mask],
               color=OBS_COLOR, alpha=0.65, s=12, linewidths=0.3,
               edgecolors=OBS_COLOR, zorder=2)
    # Influential observations — SAME SIZE as normal
    infl = ~normal_mask
    if infl.any():
        ax.scatter(obs[infl], vals[infl],
                   color=REMOVED_PT_COLOR, alpha=0.95, s=12, marker="D",
                   linewidths=0.4, edgecolors=REMOVED_PT_COLOR, zorder=3)
    # Threshold lines — amber, dashed
    THRESH_COLOR = "#B07D2A"
    if isinstance(thresh, (list, tuple)):
        for tv in thresh:
            ax.axhline(tv, color=THRESH_COLOR, linewidth=1.1, linestyle="--", zorder=1)
    else:
        ax.axhline(thresh, color=THRESH_COLOR, linewidth=1.1, linestyle="--", zorder=1)
    if zero_line:
        ax.axhline(0, color="#cccccc", linewidth=0.6, zorder=1)

    # Larger relative text so the figure stays readable at half-page scale
    ax.set_ylabel(ylabel, style="italic", fontsize=10)
    ax.set_xlabel("Observation number", style="italic", fontsize=10)
    ax.tick_params(axis="both", labelsize=11)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)

    # Per-panel threshold legend, placed above the axes (outside data area)
    th_handle = Line2D([0], [0], color=THRESH_COLOR, linewidth=1.1, linestyle="--",
                       label=thresh_label)
    leg = ax.legend(handles=[th_handle], fontsize=8, frameon=True,
                    facecolor="white", edgecolor="#cccccc",
                    loc="lower right", bbox_to_anchor=(1.0, 1.01),
                    handletextpad=0.4, borderpad=0.35)
    leg.set_zorder(5)


def _influence_figure_for_subset(sub, title, basename):
    from matplotlib.lines import Line2D
    lev, sr, dff, cks, th = _compute_influence(sub)
    obs = np.arange(1, len(sub) + 1)
    n = len(sub)
    k = 1

    # Sized for SI two-per-page: ~6.5" wide x 4.8" tall fits comfortably with title,
    # caption, and a second figure beneath it on US Letter.
    fig, axes = plt.subplots(2, 2, figsize=(6.5, 4.8))
    fig.suptitle(f"Influence diagnostics \u2014 {title}",
                 fontsize=11, color="#333333", y=1.00)

    _diag_panel(axes[0, 0], obs, lev, lev <= th["lev"], th["lev"],
                "Leverage",
                thresh_label=f"Threshold = 3 \u00d7 mean leverage = {th['lev']:.4f}")
    _diag_panel(axes[0, 1], obs, sr, np.abs(sr) < th["sr"], [-th["sr"], th["sr"]],
                "Studentized residual",
                thresh_label="Threshold = \u00b13",
                zero_line=True)
    _diag_panel(axes[1, 0], obs, dff, np.abs(dff) < th["dff"], [-th["dff"], th["dff"]],
                "DFFITS",
                thresh_label=f"Threshold = 2\u221a((k+2)/(n\u2212k\u22122)) = \u00b1{th['dff']:.3f}",
                zero_line=True)
    _diag_panel(axes[1, 1], obs, cks, cks <= th["cks"], th["cks"],
                "Cook's distance",
                thresh_label=f"Threshold = 4/n = {th['cks']:.4f}")

    legend_els = [
        Line2D([0], [0], marker="o", color="w", markerfacecolor=OBS_COLOR,
               markersize=5, label="Normal observation"),
        Line2D([0], [0], marker="D", color="w", markerfacecolor=REMOVED_PT_COLOR,
               markersize=5, label="Influential observation"),
    ]
    fig.legend(handles=legend_els, loc="lower center", ncol=2, fontsize=8,
               frameon=True, facecolor="white", edgecolor="#cccccc",
               bbox_to_anchor=(0.5, -0.04))
    plt.tight_layout(h_pad=2.3, w_pad=1.5)
    save_figure(fig, basename)
    plt.close(fig)


def figure_influence_diagnostics():
    print("Figure(s) — influence diagnostics per subtype")
    for name, (df, kind, sup) in EQUIPMENT.items():
        subtypes = list(df["Type"].unique())
        safe_eq = name.replace(" ", "_")
        if len(subtypes) <= 1:
            # single-type: one figure for the whole group
            _influence_figure_for_subset(df, name, f"Figure_influence_{safe_eq}")
        else:
            for t in subtypes:
                sub = df[df["Type"] == t].reset_index(drop=True)
                if len(sub) < 4:
                    print(f"  skip {t}: n={len(sub)} too small")
                    continue
                safe_t = t.replace(" ", "_").replace("&", "and").replace("/", "-")
                _influence_figure_for_subset(sub, t, f"Figure_influence_{safe_eq}_{safe_t}")


# ============================================================================= #
# FIGURE: POWER-LAW SCALING (BEFORE vs AFTER OUTLIER REMOVAL) PER SUBTYPE
#   For each subtype (or single-type group): a 1x2 figure
#   Left  panel = regression on full data (before outlier removal)
#   Right panel = regression after removing influential observations
#   Each panel: log-log scatter, OLS fit line, ±RMSE band, equation, R^2_adj.
#   The right panel also highlights the removed influential points in red.
#   Sized to fit two figures per A4 page.
# ============================================================================= #
def _fit_loglog(sub):
    """Degree-1 OLS log-log fit; returns intercept, slope, RMSE, R^2_adj."""
    import statsmodels.api as sm
    x = sub["logCap"].values
    y = sub["logPV"].values
    X = np.vander(x, 2, increasing=True)
    m = sm.OLS(y, X).fit()
    a, b = float(m.params[0]), float(m.params[1])
    rmse = float(np.sqrt(np.mean(m.resid**2)))
    r2_adj = float(m.rsquared_adj)
    return a, b, rmse, r2_adj


def _identify_influentials(sub, equipment_name=None):
    """Return a boolean mask of influential observations, using the EXACT
    method each notebook actually uses to build its *_clean dataframe:

      - Calciners (rotary kilns): 1.5x IQR rule on logPV only (Tukey outlier
        rule) -- see Rotary_Kilns.ipynb section 5.2. Capacity is deliberately
        not screened: an unusually small or large kiln is a legitimate part
        of the capacity range. Rotary kilns do NOT use leverage/DFFITS/Cook's
        distance at all.

      - Every other equipment (Agitated Tanks, Mixers, Filtration Units,
        Dryers): High_StudResid OR High_DFFITS OR High_Cooks, using STRICT
        ">" comparisons against the thresholds -- see each notebook's
        influence_measures() + the "influential_points = df_influence[...]"
        line. Leverage is deliberately EXCLUDED from this flag in every
        notebook (it is only used to annotate the leverage diagnostic panel,
        not to decide what gets removed).
    """
    if equipment_name == "Calciners":
        y = sub["logPV"]
        q1y, q3y = y.quantile([0.25, 0.75]); iqry = q3y - q1y
        mask = (y < q1y - 1.5 * iqry) | (y > q3y + 1.5 * iqry)
        return mask.values if hasattr(mask, "values") else mask

    import statsmodels.api as sm
    x = sub["logCap"].values
    y = sub["logPV"].values
    n = len(y); k = 1
    X = np.vander(x, k + 1, increasing=True)
    m = sm.OLS(y, X).fit()
    inf = m.get_influence()
    sr  = inf.resid_studentized_external
    dff = inf.dffits[0]
    cks = inf.cooks_distance[0]
    dff_th = 2 * np.sqrt((k + 2) / (n - k - 2)) if n - k - 2 > 0 else np.inf
    cks_th = 4 / n
    return (np.abs(sr) > 3) | (np.abs(dff) > dff_th) | (cks > cks_th)


def _draw_loglog_panel(ax, sub, a, b, rmse, r2_adj, label,
                       influentials_mask=None, removed_pts=None,
                       cap_kind="V"):
    """One log-log panel with scatter, fit line, ±RMSE band, equation box."""
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch

    x = sub["logCap"].values
    y = sub["logPV"].values

    # Smooth x grid for fit + band
    xg = np.linspace(x.min() - 0.1, x.max() + 0.1, 200)
    yg = a + b * xg

    pv_str = "P/V" if cap_kind == "V" else "P/A"
    cap_str = "V" if cap_kind == "V" else "A"

    # RMSE band
    ax.fill_between(xg, yg - rmse, yg + rmse,
                    color=RMSE_BAND_COLOR, alpha=0.22, zorder=1,
                    label=f"\u00b1RMSE ({rmse:.4f})")
    # OLS fit line
    ax.plot(xg, yg, color=FIT_LINE_COLOR, linewidth=1.4, zorder=3,
            label=f"Fit: log({pv_str}) = {a:.4f} {('+' if b>=0 else '\u2212')} {abs(b):.4f} log({cap_str})")
    # Normal observations
    if influentials_mask is not None:
        norm = ~influentials_mask
        ax.scatter(x[norm], y[norm], color=OBS_COLOR, alpha=0.70, s=12,
                   linewidths=0.3, edgecolors=OBS_COLOR, zorder=2,
                   label="Observations (log-log)")
    else:
        ax.scatter(x, y, color=OBS_COLOR, alpha=0.70, s=12,
                   linewidths=0.3, edgecolors=OBS_COLOR, zorder=2,
                   label="Observations (log-log)")
    # Removed influential points (only shown on the "after" panel; they are the
    # points that were dropped to produce the cleaned fit)
    if removed_pts is not None and len(removed_pts) > 0:
        rx, ry = removed_pts
        ax.scatter(rx, ry, color=REMOVED_PT_COLOR, alpha=0.95, s=22, marker="D",
                   linewidths=0.4, edgecolors=REMOVED_PT_COLOR, zorder=4,
                   label="Removed influential points")

    # Equation / R^2 entry (proxy handle, no visible artist)
    eq_handle = Line2D([0], [0], color="none", marker="",
                       label=f"$R^2_{{adj}}$ = {r2_adj:.4f}")

    # Compose legend in a specific order
    handles, labels = ax.get_legend_handles_labels()
    # Move RMSE band to the end
    order_keys = ["Observations", "Removed", "Fit:", "R^2", "\u00b1RMSE"]
    def sort_key(lbl):
        for i, k in enumerate(order_keys):
            if k in lbl: return i
        return 99
    handles.append(eq_handle); labels.append(eq_handle.get_label())
    paired = sorted(zip(handles, labels), key=lambda p: sort_key(p[1]))
    handles, labels = zip(*paired)

    ax.legend(handles=handles, labels=labels, fontsize=7, frameon=True,
              facecolor="white", edgecolor="#cccccc",
              loc="upper center", bbox_to_anchor=(0.5, -0.18),
              ncol=2, handletextpad=0.4, borderpad=0.4,
              labelspacing=0.3, columnspacing=1.0)

    ax.set_xlabel("log(V)", style="italic", fontsize=10)
    ax.set_ylabel("log(P/V)", style="italic", fontsize=10)
    ax.tick_params(axis="both", labelsize=9)
    # Panel letter placed to the LEFT of the panel (outside the axes)
    ax.annotate(label, xy=(-0.22, 1.02), xycoords="axes fraction",
                ha="left", va="top", fontsize=12, fontweight="bold",
                color="#333333")
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)


def _scaling_figure_for_subset(sub_all, title, basename, cap_kind, equipment_name=None):
    """Build the 1x2 before/after figure for a single subset (subtype or group)."""
    # BEFORE: fit on all observations
    a0, b0, rmse0, r2_0 = _fit_loglog(sub_all)
    inf_mask = _identify_influentials(sub_all, equipment_name=equipment_name)

    # AFTER: fit excluding influentials
    sub_clean = sub_all.loc[~inf_mask].copy()
    if len(sub_clean) < 4 or inf_mask.sum() == 0:
        # Too few points after removal, or nothing flagged — just show before
        # in the after panel as well to avoid an empty plot
        a1, b1, rmse1, r2_1 = a0, b0, rmse0, r2_0
        removed_pts = None
    else:
        a1, b1, rmse1, r2_1 = _fit_loglog(sub_clean)
        removed_pts = (sub_all.loc[inf_mask, "logCap"].values,
                       sub_all.loc[inf_mask, "logPV"].values)

    # Sized for A4 two-per-page (works for both V- and A-based fits since the
    # underlying log-log axes are equivalent). Extra height below the panels
    # makes room for the legend boxes underneath.
    fig, axes = plt.subplots(1, 2, figsize=(7.0, 4.0))

    # If using area-based equipment, relabel the axis text accordingly
    xlab = "log(V)" if cap_kind == "V" else "log(A)"
    ylab = "log(P/V)" if cap_kind == "V" else "log(P/A)"

    _draw_loglog_panel(axes[0], sub_all, a0, b0, rmse0, r2_0,
                       label="(a)",
                       cap_kind=cap_kind)
    _draw_loglog_panel(axes[1], sub_clean if len(sub_clean) >= 4 else sub_all,
                       a1, b1, rmse1, r2_1,
                       label="(b)",
                       removed_pts=removed_pts,
                       cap_kind=cap_kind)

    for ax in axes:
        ax.set_xlabel(xlab, style="italic", fontsize=10)
        ax.set_ylabel(ylab, style="italic", fontsize=10)

    fig.suptitle(f"Power-law scaling \u2014 {title}", fontsize=11,
                 color="#333333", y=1.02)
    plt.tight_layout(w_pad=3.5)
    plt.subplots_adjust(bottom=0.30, left=0.10)
    save_figure(fig, basename)
    plt.close(fig)


def figure_scaling_before_after():
    """One scaling-before/after figure per subtype (or per group for single-type)."""
    print("Figure(s) — power-law scaling before/after outlier removal")
    for name, (df, kind, sup) in EQUIPMENT.items():
        subtypes = list(df["Type"].unique())
        safe_eq = name.replace(" ", "_")
        if len(subtypes) <= 1:
            _scaling_figure_for_subset(df, name, f"Figure_scaling_{safe_eq}", kind, equipment_name=name)
        else:
            for t in subtypes:
                sub = df[df["Type"] == t].reset_index(drop=True)
                if len(sub) < 4:
                    print(f"  skip {t}: n={len(sub)} too small")
                    continue
                safe_t = t.replace(" ", "_").replace("&", "and").replace("/", "-")
                _scaling_figure_for_subset(sub, t, f"Figure_scaling_{safe_eq}_{safe_t}", kind)


# ============================================================================= #
# FIGURE: POOLED CAPACITY-CLASS BOXPLOTS (Small / Medium / Large)
#   2x2 grid covering Agitated Tanks (a), Mixers (b), Filtration Units (c),
#   Dryers (d). Calciners excluded per request. Each panel = one boxplot per
#   capacity class with the geometric mean marked, on a log-y axis.
#   Bin boundaries match the manuscript's chosen splits (volume or area, m^3/m^2).
# ============================================================================= #
# Exact row-count splits used by each notebook's Small/Medium/Large grouping
# (e.g. Mixers cell: splits_m = [57, 91, 98]). These are counts of rows -- NOT
# capacity thresholds -- applied to the capacity-sorted, outlier-removed
# dataframe. Calciners has no equivalent 3-way split (its "Grouping Analysis"
# section uses a different 2-group search, not Small/Medium/Large), so it is
# not included here and keeps its existing "All kilns" single-box treatment.
SPLITS = {
    "Agitated Tanks":   [35, 54, 58],
    "Mixers":           [57, 91, 98],
    "Filtration Units": [43, 62, 40],
    "Dryers":           [31, 33, 39],
}


def _clean_by_subtype(df, equipment_name=None):
    """Remove influential points using each subtype's OWN regression fit,
    exactly matching how every notebook builds its *_clean dataframe: outliers
    are identified per-subtype, but the removal is applied against the whole
    capacity-sorted frame -- so the result stays sorted by capacity across all
    subtypes combined (required for the row-count Small/Medium/Large splits
    to line up with the right capacity ranges)."""
    subtypes = df["Type"].unique()
    if len(subtypes) <= 1:
        mask = _identify_influentials(df, equipment_name=equipment_name)
        return df.loc[~np.asarray(mask)].copy()
    remove_idx = []
    for t in subtypes:
        sub = df[df["Type"] == t]
        if len(sub) < 4:
            continue
        mask = _identify_influentials(sub, equipment_name=equipment_name)
        remove_idx.extend(sub.index[np.asarray(mask)].tolist())
    return df.loc[~df.index.isin(remove_idx)].copy()


def figure_pooled_capacity_boxplots():
    print("Figure — pooled capacity-class boxplots (Tanks, Mixers, Filters, Dryers)")
    from scipy.stats import gmean
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch

    panels = [
        ("Agitated Tanks", "Agitated Tanks", [
            ("Small",  0.005, 0.20,   "[0.005, 0.20)"),
            ("Medium", 0.20,  2.00,   "[0.20, 2.00)"),
            ("Large",  2.00,  135.40, "[2.00, 135.40)"),
        ], "Volume [m\u00b3]"),
        ("Mixers", "Mixers", [
            ("Small",  0.004, 0.2,  "[0.004, 0.2)"),
            ("Medium", 0.2,   1.4,  "[0.2, 1.4)"),
            ("Large",  1.4,   20,   "[1.4, 20)"),
        ], "Volume [m\u00b3]"),
        ("Filtration Units", "Filtration Units", [
            ("Small",  0.06, 2.5,  "[0.06, 2.5)"),
            ("Medium", 2.5,  80,   "[2.5, 80)"),
            ("Large",  80,   1000, "[80, 1000)"),
        ], "Area [m\u00b2]"),
        ("Dryers", "Dryers", [
            ("Small",  0.75, 14,  "[0.75, 14)"),
            ("Medium", 14,   42,  "[14, 42)"),
            ("Large",  42,   180, "[42, 180)"),
        ], "Area [m\u00b2]"),
        ("Rotary Kilns", "Calciners", [
            ("All", 0, 1e9, "All kilns"),
        ], "Volume [m\u00b3]"),
    ]
    letters = ["a", "b", "c", "d", "e"]

    # Border colors — darker shade of the fill color per class
    CAP_BORDER = {"Small": "#0E1F33", "Medium": "#1F3A22", "Large": "#3F0D17"}

    # A4 portrait is 8.27" x 11.69" usable area ~7" wide.
    # Half-page (top half): width ~7", height ~5.2" comfortably fits a 2x2 boxplot grid
    # with readable labels and a bottom legend.
    fig, axes = plt.subplots(2, 3, figsize=(11, 6.8))
    axes[1, 2].set_visible(False)  # hide unused 6th cell

    for idx, (name, key, bins, cap_unit_label) in enumerate(panels):
        ax = axes.flatten()[idx]
        df, kind, sup = EQUIPMENT[key]
        unit = unit_label(sup)

        # ── Outlier removal (per-subtype, exact notebook method) ──────────
        df = _clean_by_subtype(df, equipment_name=key)

        groups, bin_labels_used, bin_classes = [], [], []
        if key in SPLITS:
            # Row-count splits, exactly matching e.g. "splits_m = [57, 91, 98]"
            # applied to the capacity-sorted, cleaned dataframe.
            splits = SPLITS[key]
            indices = np.cumsum([0] + splits).tolist()
            for i, (cls, lo, hi, lbl) in enumerate(bins):
                start, end = indices[i], indices[i + 1]
                vals = df["PV"].values[start:end]
                vals = vals[vals > 0]
                groups.append(vals)
                bin_labels_used.append(lbl)
                bin_classes.append(cls)
        else:
            # Calciners: no Small/Medium/Large split in the notebook -- single
            # "All kilns" group over the whole (IQR-)cleaned dataframe.
            for (cls, lo, hi, lbl) in bins:
                vals = df["PV"].values
                vals = vals[vals > 0]
                groups.append(vals)
                bin_labels_used.append(lbl)
                bin_classes.append(cls)

        # Map each bin to its color; "All" uses the Medium palette
        cls_face_map   = {"Small": CAP_FACE["Small"],   "Medium": CAP_FACE["Medium"],
                          "Large": CAP_FACE["Large"],   "All": CAP_FACE["Medium"]}
        cls_border_map = {"Small": CAP_BORDER["Small"], "Medium": CAP_BORDER["Medium"],
                          "Large": CAP_BORDER["Large"], "All": CAP_BORDER["Medium"]}
        cls_alpha_map  = {"Small": CAP_ALPHA["Small"],  "Medium": CAP_ALPHA["Medium"],
                          "Large": CAP_ALPHA["Large"],  "All": CAP_ALPHA["Medium"]}
        cap_face   = [cls_face_map[c]   for c in bin_classes]
        cap_border = [cls_border_map[c] for c in bin_classes]

        bp = ax.boxplot(
            groups, patch_artist=True, widths=0.55,
            medianprops=dict(color="#222222", linewidth=1.5),
            whiskerprops=dict(linewidth=0.8, linestyle="--", color="#777777"),
            capprops=dict(linewidth=0.8, color="#777777"),
            flierprops=dict(marker="o", markersize=3, alpha=0.45,
                            markerfacecolor="#888888", markeredgecolor="#888888",
                            linewidth=0),
        )
        # Apply fill + darker border per class
        for patch, fc, bc, cls in zip(bp["boxes"], cap_face, cap_border, bin_classes):
            patch.set_facecolor(fc)
            patch.set_alpha(cls_alpha_map[cls])
            patch.set_edgecolor(bc)
            patch.set_linewidth(1.4)

        # GM markers + labels — text placed OUTSIDE each box (to the right) so it
        # never overlaps the box interior
        for i, vals in enumerate(groups, start=1):
            if len(vals) == 0:
                continue
            gm = gmean(vals)
            ax.scatter(i, gm, color=GM_MARKER_COLOR, marker="^", s=46, zorder=4,
                       edgecolors=GM_MARKER_COLOR, linewidths=0.5)
            # Place label clearly to the right of the box (~half-width offset),
            # vertically aligned with the GM
            ax.annotate(
                f"{gm:.2f}",
                xy=(i + 0.30, gm),
                ha="left", va="center",
                fontsize=8.5, color=GM_MARKER_COLOR,
                bbox=dict(facecolor="white", edgecolor="none", pad=1.5),
                zorder=5,
            )

        ax.set_yscale("log")
        ax.set_xticks(list(range(1, len(groups) + 1)))
        ax.set_xticklabels(bin_labels_used, fontsize=9)
        ax.set_xlabel(f"{name} {cap_unit_label}", style="italic", fontsize=10)
        ax.set_ylabel(f"Power intensity [{unit}]", style="italic", fontsize=10)
        ax.tick_params(axis="both", labelsize=9)
        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)

        # Extend x-axis slightly so the GM label has room and the right-most box
        # isn't crammed against the spine
        ax.set_xlim(0.4, len(groups) + 0.85)

        ax.annotate(f"({letters[idx]})", xy=(-0.18, 1.05), xycoords="axes fraction",
                    ha="left", va="top", fontsize=12, fontweight="bold",
                    color="#333333")

    # Shared legend below the grid
    legend_els = [
        Patch(facecolor=CAP_FACE["Small"],  alpha=CAP_ALPHA["Small"],
              edgecolor=CAP_BORDER["Small"],  linewidth=1.2, label="Small"),
        Patch(facecolor=CAP_FACE["Medium"], alpha=CAP_ALPHA["Medium"],
              edgecolor=CAP_BORDER["Medium"], linewidth=1.2, label="Medium"),
        Patch(facecolor=CAP_FACE["Large"],  alpha=CAP_ALPHA["Large"],
              edgecolor=CAP_BORDER["Large"],  linewidth=1.2, label="Large"),
        Line2D([0], [0], marker="^", color="w", markerfacecolor=GM_MARKER_COLOR,
               markersize=8, label="Geometric mean"),
    ]
    fig.legend(handles=legend_els, loc="lower center", ncol=4, fontsize=9,
               frameon=True, facecolor="white", edgecolor="#cccccc",
               bbox_to_anchor=(0.5, -0.01))

    plt.tight_layout(h_pad=2.5, w_pad=2.5)
    plt.subplots_adjust(bottom=0.10, left=0.10)
    save_figure(fig, "Figure_pooled_capacity_boxplots")
    plt.close(fig)


# ============================================================================= #
# FIGURE: SUBTYPE-GROUPED CAPACITY-CLASS BOXPLOTS, ONE FIGURE PER EQUIPMENT TYPE
#   Same Small/Medium/Large color/border scheme as the pooled figure, but with
#   subtypes laid out along the x-axis and three boxes per subtype (Small,
#   Medium, Large). Only generated for multi-subtype equipment: Mixers,
#   Filtration Units, Dryers. Tanks and Calciners (single-type) are skipped.
# ============================================================================= #
def figure_subtype_capacity_boxplots():
    print("Figure(s) — subtype-grouped capacity-class boxplots")
    from scipy.stats import gmean
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch

    CAP_BORDER = {"Small": "#0E1F33", "Medium": "#1F3A22", "Large": "#3F0D17"}
    class_labels = ["Small", "Medium", "Large"]

    # Same physical thresholds as the pooled figure
    BIN_CONFIG = {
        "Mixers":           [(0.004, 0.2,  "Small"),
                             (0.2,   1.4,  "Medium"),
                             (1.4,   20,   "Large")],
        "Filtration Units": [(0.06,  2.5,  "Small"),
                             (2.5,   80,   "Medium"),
                             (80,    1000, "Large")],
        "Dryers":           [(0.75,  14,   "Small"),
                             (14,    42,   "Medium"),
                             (42,    180,  "Large")],
    }
    CAP_UNIT_LABEL = {
        "Mixers":           "Volume [m\u00b3]",
        "Filtration Units": "Area [m\u00b2]",
        "Dryers":           "Area [m\u00b2]",
    }

    # Per-equipment layout tuning so each reads well on A4
    LAYOUT = {
        "Mixers": {           # 6 subtypes — tight but readable
            "figsize": (8.0, 4.8),
            "intra_gap": 0.55, "inter_gap": 0.85, "pooled_gap": 1.35,
            "box_width": 0.42, "label_dx": 0.22,
            "label_fs": 7, "xtick_fs": 8, "wrap": 10,
        },
        "Filtration Units": { # 3 subtypes — was too spread, make compact + bigger text
            "figsize": (6.4, 4.8),
            "intra_gap": 0.70, "inter_gap": 1.20, "pooled_gap": 1.70,
            "box_width": 0.52, "label_dx": 0.28,
            "label_fs": 9, "xtick_fs": 10, "wrap": 14,
        },
        "Dryers": {           # 5 subtypes — more inter-gap so Rotary Drum / Rotary Plate separate
            "figsize": (8.0, 4.8),
            "intra_gap": 0.58, "inter_gap": 1.20, "pooled_gap": 1.70,
            "box_width": 0.44, "label_dx": 0.24,
            "label_fs": 7.5, "xtick_fs": 9, "wrap": 12,
        },
    }

    for eq_name, bins in BIN_CONFIG.items():
        df, kind, sup = EQUIPMENT[eq_name]
        # ── Outlier removal (per-subtype, exact notebook method) ──────────
        df = _clean_by_subtype(df, equipment_name=eq_name)
        # ── Size class assigned by GLOBAL row-count position (matches the
        #    notebook: splits are applied to the whole capacity-sorted,
        #    cleaned frame, not separately within each subtype) ───────────
        splits = SPLITS[eq_name]
        indices = np.cumsum([0] + splits).tolist()
        size_class = np.empty(len(df), dtype=object)
        for i, cls in enumerate(class_labels):
            size_class[indices[i]:indices[i + 1]] = cls
        df = df.copy()
        df["SizeClass"] = size_class
        unit = unit_label(sup)
        subtypes = list(df["Type"].unique())
        n_sub = len(subtypes)
        L = LAYOUT[eq_name]

        # Rightmost group is the pooled (all-subtypes-combined) model.
        # POOLED_KEY is a sentinel, not a real "Type" value.
        POOLED_KEY = "__POOLED__"
        groups_order = subtypes + [POOLED_KEY]

        fig, ax = plt.subplots(figsize=L["figsize"])

        intra_gap  = L["intra_gap"]
        inter_gap  = L["inter_gap"]
        pooled_gap = L["pooled_gap"]
        box_width  = L["box_width"]
        label_dx   = L["label_dx"]

        positions = []
        center_positions = []
        x = 1.0
        for s_idx, item in enumerate(groups_order):
            trio_positions = [x, x + intra_gap, x + 2 * intra_gap]
            positions.append(trio_positions)
            center_positions.append(trio_positions[1])
            gap = pooled_gap if item == subtypes[-1] else inter_gap
            x = trio_positions[-1] + gap

        pooled_x_start = positions[-1][0] - intra_gap  # left edge of pooled trio's slot

        for s_idx, (item, trio_pos) in enumerate(zip(groups_order, positions)):
            sub = df if item == POOLED_KEY else df[df["Type"] == item]
            for cls_idx, ((lo, hi, cls), pos) in enumerate(zip(bins, trio_pos)):
                vals = sub.loc[sub["SizeClass"] == cls, "PV"].values
                vals = vals[vals > 0]
                if len(vals) == 0:
                    continue

                bp = ax.boxplot(
                    [vals], positions=[pos], widths=box_width,
                    patch_artist=True,
                    medianprops=dict(color="#222222", linewidth=1.2),
                    whiskerprops=dict(linewidth=0.6, linestyle="--", color="#777777"),
                    capprops=dict(linewidth=0.6, color="#777777"),
                    flierprops=dict(marker="o", markersize=2.2, alpha=0.45,
                                    markerfacecolor="#888888",
                                    markeredgecolor="#888888", linewidth=0),
                )
                patch = bp["boxes"][0]
                patch.set_facecolor(CAP_FACE[cls])
                patch.set_alpha(CAP_ALPHA[cls])
                patch.set_edgecolor(CAP_BORDER[cls])
                patch.set_linewidth(1.1)

                gm = gmean(vals)
                ax.scatter(pos, gm, color=GM_MARKER_COLOR, marker="^", s=28,
                           zorder=4, edgecolors=GM_MARKER_COLOR, linewidths=0.4)
                # Alternate label side: Small→right, Medium→left, Large→right
                if cls == "Medium":
                    lbl_x = pos - label_dx
                    lbl_ha = "right"
                else:
                    lbl_x = pos + label_dx
                    lbl_ha = "left"
                ax.annotate(
                    f"{gm:.2f}",
                    xy=(lbl_x, gm),
                    ha=lbl_ha, va="center",
                    fontsize=L["label_fs"], color=GM_MARKER_COLOR,
                    bbox=dict(facecolor="white", edgecolor="none", pad=0.8),
                    zorder=5,
                )

        ax.set_yscale("log")
        ax.set_xticks(center_positions)
        import textwrap
        wrapped_labels = [
            textwrap.fill(s.replace(" Mixers", "").replace(" Filters", "")
                            .replace(" Dryers", ""), L["wrap"])
            for s in subtypes
        ] + ["Pooled\n(all sub-types)"]
        ax.set_xticklabels(wrapped_labels, fontsize=L["xtick_fs"])
        # Bold the "Pooled" tick label so it reads as distinct from subtypes
        ax.get_xticklabels()[-1].set_fontweight("bold")
        # Vertical separator between the last subtype and the pooled group
        sep_x = (positions[-2][-1] + positions[-1][0]) / 2
        ax.axvline(sep_x, color="#999999", linewidth=0.9, linestyle=":",
                   zorder=0)
        ax.set_xlabel(f"{eq_name} sub-types — {CAP_UNIT_LABEL[eq_name]}",
                      style="italic", fontsize=10)
        ax.set_ylabel(f"Power intensity [{unit}]", style="italic", fontsize=10)
        ax.tick_params(axis="both", labelsize=L["xtick_fs"])
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.set_xlim(0.4, positions[-1][-1] + label_dx + 0.7)

        legend_els = [
            Patch(facecolor=CAP_FACE["Small"],  alpha=CAP_ALPHA["Small"],
                  edgecolor=CAP_BORDER["Small"],  linewidth=1.2, label="Small"),
            Patch(facecolor=CAP_FACE["Medium"], alpha=CAP_ALPHA["Medium"],
                  edgecolor=CAP_BORDER["Medium"], linewidth=1.2, label="Medium"),
            Patch(facecolor=CAP_FACE["Large"],  alpha=CAP_ALPHA["Large"],
                  edgecolor=CAP_BORDER["Large"],  linewidth=1.2, label="Large"),
            Line2D([0], [0], marker="^", color="w",
                   markerfacecolor=GM_MARKER_COLOR, markersize=7,
                   label="Geometric mean"),
        ]
        ax.legend(handles=legend_els, loc="upper center",
                  bbox_to_anchor=(0.5, -0.18),
                  ncol=4, fontsize=8.5, frameon=True, facecolor="white",
                  edgecolor="#cccccc", handletextpad=0.4, borderpad=0.5,
                  columnspacing=1.5)

        plt.tight_layout()
        plt.subplots_adjust(bottom=0.22)
        safe_eq = eq_name.replace(" ", "_")
        save_figure(fig, f"Figure_subtype_boxplots_{safe_eq}")
        plt.close(fig)


# ============================================================================= #
# FIGURE: ASSUMPTION CHECK (Residuals vs Fitted + Normal Q–Q)
#   5 figures total — one per equipment type. Multi-subtype equipment shows
#   all subtypes as rows plus a pooled row at the bottom. Single-type
#   equipment shows just one row.
#   Each row: (left) residuals vs fitted, (right) Normal Q–Q.
# ============================================================================= #
def figure_assumption_checks():
    print("Figure(s) — assumption checks (residuals vs fitted + Q-Q)")
    import statsmodels.api as sm
    from scipy import stats as sp_stats

    def _draw_row(ax_resid, ax_qq, sub, row_label, equipment_name=None):
        """Fit OLS after outlier removal, draw one row of panels."""
        inf_mask = _identify_influentials(sub, equipment_name=equipment_name)
        sub_clean = sub.loc[~inf_mask].reset_index(drop=True)
        if len(sub_clean) < 4 or inf_mask.sum() == 0:
            sub_clean = sub

        x = sub_clean["logCap"].values
        y = sub_clean["logPV"].values
        X = np.vander(x, 2, increasing=True)
        model = sm.OLS(y, X).fit()
        fitted = model.fittedvalues
        resid = model.resid

        # (left) Residuals vs Fitted
        ax_resid.scatter(fitted, resid, color=OBS_COLOR, alpha=0.65, s=12,
                         linewidths=0.3, edgecolors=OBS_COLOR, zorder=2)
        ax_resid.axhline(0, color=REF_LINE_COLOR, linewidth=1.2, linestyle="--", zorder=1)
        ax_resid.set_xlabel("Fitted values (log scale)", style="italic", fontsize=8.5)
        ax_resid.set_ylabel("Residuals", style="italic", fontsize=8.5)
        ax_resid.tick_params(axis="both", labelsize=8)
        ax_resid.spines["top"].set_visible(False)
        ax_resid.spines["right"].set_visible(False)

        # (right) Normal Q–Q
        resid_sorted = np.sort(resid)
        n = len(resid_sorted)
        theoretical_q = sp_stats.norm.ppf((np.arange(1, n + 1) - 0.5) / n)
        ax_qq.scatter(theoretical_q, resid_sorted, color=OBS_COLOR, alpha=0.65, s=12,
                      linewidths=0.3, edgecolors=OBS_COLOR, zorder=2)
        q25_t, q75_t = np.percentile(theoretical_q, [25, 75])
        q25_s, q75_s = np.percentile(resid_sorted, [25, 75])
        slope = (q75_s - q25_s) / (q75_t - q25_t) if q75_t != q25_t else 1
        intercept = q25_s - slope * q25_t
        xline = np.array([theoretical_q.min() - 0.3, theoretical_q.max() + 0.3])
        ax_qq.plot(xline, intercept + slope * xline,
                   color=REF_LINE_COLOR, linewidth=1.3, zorder=1)
        ax_qq.set_xlabel("Theoretical quantiles", style="italic", fontsize=8.5)
        ax_qq.set_ylabel("Sample quantiles", style="italic", fontsize=8.5)
        ax_qq.tick_params(axis="both", labelsize=8)
        ax_qq.spines["top"].set_visible(False)
        ax_qq.spines["right"].set_visible(False)

        # Row label on the left side
        ax_resid.annotate(
            row_label, xy=(-0.32, 0.5), xycoords="axes fraction",
            ha="center", va="center", rotation=90,
            fontsize=9, color="#333333", style="italic")

    # Max rows per page to keep panels readable on A4
    MAX_ROWS_PER_PAGE = 4

    for eq_name, (df, kind, sup) in EQUIPMENT.items():
        subtypes = list(df["Type"].unique())
        safe_eq = eq_name.replace(" ", "_")

        # Build list of (subset, label) rows
        rows = []
        if len(subtypes) > 1:
            for t in subtypes:
                sub = df[df["Type"] == t].reset_index(drop=True)
                if len(sub) >= 4:
                    short = (t.replace(" Mixers", "").replace(" Filters", "")
                              .replace(" Dryers", ""))
                    rows.append((sub, short))
            rows.append((df, f"{eq_name}\n(pooled)"))
        else:
            rows.append((df, eq_name))

        # Split into pages if too many rows
        if len(rows) <= MAX_ROWS_PER_PAGE:
            pages = [rows]
        else:
            mid = (len(rows) + 1) // 2  # first page gets slightly more if odd
            pages = [rows[:mid], rows[mid:]]

        for page_idx, page_rows in enumerate(pages):
            n_rows = len(page_rows)
            row_h = 2.8
            fig_h = n_rows * row_h + 0.8
            fig, axes = plt.subplots(n_rows, 2, figsize=(7.5, fig_h))
            if n_rows == 1:
                axes = axes.reshape(1, 2)

            suffix = ""
            if len(pages) > 1:
                suffix = f" (part {page_idx + 1})"

            fig.suptitle(f"Assumption checks \u2014 {eq_name}{suffix}",
                         fontsize=12, color="#333333", y=0.995)

            # Column headers
            axes[0, 0].set_title("Residuals vs fitted", fontsize=9.5,
                                 color="#555555", style="italic", pad=6)
            axes[0, 1].set_title("Normal Q\u2013Q", fontsize=9.5,
                                 color="#555555", style="italic", pad=6)

            for r, (sub, label) in enumerate(page_rows):
                _draw_row(axes[r, 0], axes[r, 1], sub, label, equipment_name=eq_name)

            plt.tight_layout(h_pad=2.0, w_pad=2.0)
            plt.subplots_adjust(left=0.16)

            if len(pages) == 1:
                save_figure(fig, f"Figure_assumptions_{safe_eq}")
            else:
                save_figure(fig, f"Figure_assumptions_{safe_eq}_part{page_idx + 1}")
            plt.close(fig)


# ============================================================================= #
# FIGURE: MULTI-SUBTYPE SCALING OVERLAY — SINGLE COMBINED 1×3 FIGURE
#   Three panels side by side: (a) Mixers, (b) Filtration, (c) Dryers.
#   Each panel: raw/power scale (log-log axes) with scatter + OLS fit lines
#   per subtype (after outlier removal). Equations in color-coded boxes placed
#   clear of all data and lines.
# ============================================================================= #
def figure_multi_subtype_scaling():
    print("Figure — combined multi-subtype scaling (1×3)")
    import statsmodels.api as sm
    from matplotlib.lines import Line2D

    # Equipment to include (multi-subtype only)
    panel_config = [
        ("Mixers",           "V"),
        ("Filtration Units", "A"),
        ("Dryers",           "A"),
    ]
    letters = ["a", "b", "c"]

    fig, axes = plt.subplots(1, 3, figsize=(17, 9.0))

    for pidx, (eq_name, kind) in enumerate(panel_config):
        ax = axes[pidx]
        df, _, sup = EQUIPMENT[eq_name]
        subtypes = list(df["Type"].unique())
        unit = unit_label(sup)
        cap_word = "Volume" if kind == "V" else "Area"
        cap_axis = "V" if kind == "V" else "A"
        pv_str = "P/V" if kind == "V" else "P/A"

        legend_handles = []
        eq_texts = []  # collect (equation_string, color) for the box

        for t in subtypes:
            st = SUBTYPE_STYLES.get(t, {"color": "#333333", "marker": "o",
                                         "linestyle": "-"})
            sub = df[df["Type"] == t].reset_index(drop=True)
            if len(sub) < 4:
                continue

            # Remove influential points
            inf_mask = _identify_influentials(sub)
            sub_clean = sub.loc[~inf_mask].reset_index(drop=True)
            if len(sub_clean) < 4:
                sub_clean = sub

            # Fit OLS on clean data
            x = sub_clean["logCap"].values
            y = sub_clean["logPV"].values
            X = np.vander(x, 2, increasing=True)
            model = sm.OLS(y, X).fit()
            a, b = float(model.params[0]), float(model.params[1])

            # Scatter (after outlier removal)
            ax.scatter(sub_clean["Capacity"], sub_clean["PV"],
                       color=st["color"], marker=st["marker"],
                       s=14, alpha=0.55, edgecolors=st["color"],
                       linewidths=0.3, zorder=2)

            # Fit line in raw/power scale
            xg = np.linspace(sub_clean["logCap"].min() ,
                             sub_clean["logCap"].max() , 200)
            yg = a + b * xg
            ax.plot(10**xg, 10**yg, color=st["color"], linewidth=1.5,
                    linestyle=st["linestyle"], zorder=3)

            # Build equation string
            sign = "+" if b >= 0 else "\u2212"
            short_name = (t.replace(" Mixers", "").replace(" Filters", "")
                           .replace(" Dryers", ""))
            eq_str = f"{short_name}: log({pv_str}) = {a:.2f} {sign} {abs(b):.2f} log({cap_axis})"
            eq_texts.append((eq_str, st["color"]))

            # Legend handle
            legend_handles.append(
                Line2D([0], [0], marker=st["marker"], color=st["color"],
                       markerfacecolor=st["color"], markeredgecolor=st["color"],
                       markersize=5, linewidth=1.2, linestyle=st["linestyle"],
                       label=short_name))

        # --- Equations below axes: one per line, single column -----------------
        for ei, (eq_str, eq_col) in enumerate(eq_texts):
            y_frac = -0.28 - ei * 0.115          # one row per equation, top-down
            ax.annotate(
                eq_str, xy=(0.01, y_frac), xycoords="axes fraction",
                fontsize=12, color=eq_col, fontweight="bold",
                ha="left", va="top",
                bbox=dict(facecolor="white", edgecolor=eq_col,
                          alpha=0.90, pad=2.0, linewidth=0.8),
                annotation_clip=False, zorder=6)

        ax.set_xscale("log"); ax.set_yscale("log")
        ax.set_xlabel(f"log({cap_axis})",
                      style="italic", fontsize=13)
        ax.set_ylabel(f"log({pv_str})", style="italic", fontsize=14)
        ax.tick_params(axis="both", labelsize=12)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

        # Panel letter — left of each panel
        ax.annotate(f"({letters[pidx]})", xy=(-0.15, 1.04),
                    xycoords="axes fraction", ha="left", va="bottom",
                    fontsize=16, fontweight="bold", color="#333333")

        # Legend floated just above the axes — never overlaps data or fit lines
        ax.legend(handles=legend_handles, fontsize=12, frameon=True,
                  facecolor="white", edgecolor="#cccccc", framealpha=1.0,
                  loc="lower right", bbox_to_anchor=(1.0, 1.02),
                  handletextpad=0.4, borderpad=0.5,
                  labelspacing=0.35, markerscale=1.1)

    # Single-col equations: bottom=0.52 fits 6 rows; no left override → tight_layout
    # handles y-label width automatically, giving subplots proper landscape aspect.
    plt.tight_layout(w_pad=1.5)
    plt.subplots_adjust(bottom=0.52, top=0.82)
    save_figure(fig, "Figure_multitype_scaling_combined")
    plt.close(fig)


# ============================================================================= #
# MAIN  —  call the figure builders you want to (re)generate
# ============================================================================= #

# ============================================================================= #
# FIGURE: POOLED POWER-LAW SCALING — ALL EQUIPMENT TYPES (after outlier removal)
#   2×3 grid; one panel per equipment type; pooled (not subtype-split) model.
#   Style mirrors the reference figure: navy obs, red-diamond removals,
#   red OLS line, sage ±RMSE band, equation+R² box lower-left.
# ============================================================================= #
def _draw_pooled_scaling_panel(ax, eq_name, kind, sup, letter, show_outliers=True):
    """Draw one panel of the pooled-scaling figure: scatter, OLS fit line,
    +/-RMSE band, and equation box. Rotary kilns (EQUIPMENT key "Calciners")
    take a constant, size-independent model in place of a power law; every
    other element of the panel is drawn identically.

    Outlier removal is done PER-SUBTYPE (via _clean_by_subtype), exactly
    matching how every notebook builds its pooled *_clean dataframe: each
    subtype gets its own regression fit to flag outliers, and only then are
    the results combined into one pooled, capacity-sorted frame. Fitting a
    single regression across all subtypes pooled together (as an outlier
    detection step) does NOT match the notebooks and gives a different
    equation -- this was the earlier bug.
    """
    import statsmodels.api as sm

    df, _, _ = EQUIPMENT[eq_name]

    pv_str  = "P/V" if kind == "V" else "P/A"
    cap_str = "V"   if kind == "V" else "A"
    exp_cap = "3" if kind == "V" else "2"
    x_unit  = f"m$^{{{exp_cap}}}$"
    y_unit  = f"kW\u00b7m$^{{{sup}}}$"

    if eq_name == "Calciners":
        # ── Rotary kilns: constant (size-independent) model ──────────────
        #    Power intensity does not scale with capacity, so the fitted
        #    object is the geometric mean in log space rather than a power
        #    law. The band is +/-RMSE about that constant, computed the same
        #    way as for the regression panels: sqrt(mean(resid**2)), which for
        #    a constant model is the log-space SD about the mean.
        #    Source: Rotary_Kilns.ipynb, section 5.3.4 ("generating Figure 1e").
        mask  = np.asarray(_identify_influentials(df, equipment_name=eq_name))
        clean = df.loc[~mask]

        x_c = clean["logCap"].values
        y_c = clean["logPV"].values

        mean_log = float(y_c.mean())
        rmse     = float(np.sqrt(np.mean((y_c - mean_log) ** 2)))

        # ── +/-RMSE band + constant model line ─────────────────────
        xg = np.linspace(x_c.min() - 0.15, x_c.max() + 0.15, 300)
        ax.fill_between(xg, mean_log - rmse, mean_log + rmse,
                        color=RMSE_BAND_COLOR, alpha=0.22, zorder=1)
        ax.plot(xg, np.full_like(xg, mean_log),
                color=FIT_LINE_COLOR, linewidth=1.5, zorder=3)

        # ── Scatter: clean observations ─────────────────────────
        ax.scatter(x_c, y_c, color=OBS_COLOR, alpha=0.70, s=14,
                   linewidths=0.3, edgecolors=OBS_COLOR, zorder=2)

        # ── Scatter: removed points (only if show_outliers) ──────────
        if show_outliers:
            removed = df.loc[mask]
            if len(removed) > 0:
                ax.scatter(removed["logCap"].values, removed["logPV"].values,
                           color=REMOVED_PT_COLOR, alpha=0.95, s=26, marker="D",
                           linewidths=0.4, edgecolors=REMOVED_PT_COLOR, zorder=4)

        # ── +/-RMSE annotation + model box ──────────────────────
        ax.text(0.97, 0.97, f"\u00b1RMSE = {rmse:.4f}",
                transform=ax.transAxes, fontsize=11,
                va="top", ha="right", color=RMSE_BAND_COLOR,
                bbox=dict(facecolor="white", edgecolor=RMSE_BAND_COLOR,
                          alpha=0.92, boxstyle="round,pad=0.35", linewidth=0.8))

        ax.text(0.03, 0.04, f"log({pv_str}) = {mean_log:.4f}",
                transform=ax.transAxes, fontsize=11,
                va="bottom", ha="left",
                bbox=dict(facecolor="white", edgecolor="#aaaaaa",
                          alpha=0.92, boxstyle="round,pad=0.4", linewidth=0.8))
    else:
        # ── Identify & remove influentials PER-SUBTYPE, exactly matching how
        #    each notebook builds its pooled *_clean dataframe ─────────────
        subtypes = df["Type"].unique()
        if len(subtypes) <= 1:
            mask = np.asarray(_identify_influentials(df, equipment_name=eq_name))
        else:
            remove_idx = []
            for t in subtypes:
                sub = df[df["Type"] == t]
                if len(sub) < 4:
                    continue
                m = _identify_influentials(sub, equipment_name=eq_name)
                remove_idx.extend(sub.index[np.asarray(m)].tolist())
            mask = df.index.isin(remove_idx)

        sub_clean = df.loc[~mask].copy().reset_index(drop=True)
        if len(sub_clean) < 4:
            sub_clean = df.copy()
            mask = np.zeros(len(df), dtype=bool)

        # ── Fit on clean data ────────────────────────────────────────────
        x_c = sub_clean["logCap"].values
        y_c = sub_clean["logPV"].values
        X   = np.vander(x_c, 2, increasing=True)
        mdl = sm.OLS(y_c, X).fit(cov_type="HC3")
        a, b   = float(mdl.params[0]), float(mdl.params[1])
        rmse   = float(np.sqrt(np.mean(mdl.resid**2)))
        r2_adj = float(mdl.rsquared_adj)

        # ── Scatter: clean observations ──────────────────────────────────
        ax.scatter(x_c, y_c, color=OBS_COLOR, alpha=0.70, s=14,
                   linewidths=0.3, edgecolors=OBS_COLOR, zorder=2)

        # ── Scatter: removed influential points (only if show_outliers) ──
        if show_outliers:
            removed = df.loc[mask]
            if len(removed) > 0:
                ax.scatter(removed["logCap"].values, removed["logPV"].values,
                           color=REMOVED_PT_COLOR, alpha=0.95, s=26, marker="D",
                           linewidths=0.4, edgecolors=REMOVED_PT_COLOR, zorder=4)

        # ── +/-RMSE band + OLS fit line ───────────────────────────────────
        xg = np.linspace(x_c.min() - 0.15, x_c.max() + 0.15, 300)
        yg = a + b * xg
        ax.fill_between(xg, yg - rmse, yg + rmse,
                        color=RMSE_BAND_COLOR, alpha=0.22, zorder=1)
        ax.plot(xg, yg, color=FIT_LINE_COLOR, linewidth=1.5, zorder=3)

        # ── +/-RMSE annotation + equation box ─────────────────────────────
        ax.text(0.97, 0.97, f"\u00b1RMSE = {rmse:.4f}",
                transform=ax.transAxes, fontsize=11,
                va="top", ha="right", color=RMSE_BAND_COLOR,
                bbox=dict(facecolor="white", edgecolor=RMSE_BAND_COLOR,
                          alpha=0.92, boxstyle="round,pad=0.35", linewidth=0.8))

        sign    = "+" if b >= 0 else "\u2212"
        eq_line = (f"log({pv_str}) = {a:.4f} {sign} {abs(b):.4f} log({cap_str})\n"
                   f"$R^2_{{adj}}$ = {r2_adj:.4f}")
        ax.text(0.03, 0.04, eq_line,
                transform=ax.transAxes, fontsize=11,
                va="bottom", ha="left",
                bbox=dict(facecolor="white", edgecolor="#aaaaaa",
                          alpha=0.92, boxstyle="round,pad=0.4", linewidth=0.8))

    # ── Axis labels & cosmetics (applies to every panel, Calciners too) ────
    ax.set_xlabel(f"log({cap_str})", style="italic", fontsize=12)
    ax.set_ylabel(f"log({pv_str})",  style="italic", fontsize=12)
    ax.tick_params(axis="both", labelsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    # ── Panel letter ────────────────────────────────────────────────────────
    ax.annotate(f"({letter})", xy=(-0.14, 1.04),
                xycoords="axes fraction", ha="left", va="bottom",
                fontsize=14, fontweight="bold", color="#333333")


def figure_pooled_scaling_all():
    """Pooled log-log fit (after outlier removal) for every equipment type,
    WITH removed/influential points marked in red. Unified legend in cell 6;
    per-panel +/-RMSE annotation at top of each axes."""
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch

    print("Figure — pooled power-law scaling, all equipment types (with outliers marked)")

    equipment_order = [
        ("Agitated Tanks",   "V", "-3"),
        ("Mixers",           "V", "-3"),
        ("Filtration Units", "A", "-2"),
        ("Dryers",           "A", "-2"),
        ("Calciners",        "V", "-3"),
    ]
    letters = ["a", "b", "c", "d", "e"]

    fig, axes = plt.subplots(2, 3, figsize=(14, 9))

    for pidx, (eq_name, kind, sup) in enumerate(equipment_order):
        row, col = divmod(pidx, 3)
        ax = axes[row, col]
        _draw_pooled_scaling_panel(ax, eq_name, kind, sup, letters[pidx],
                                   show_outliers=True)

    # ── Unified legend in cell (1, 2) ───────────────────────────────────────
    ax6 = axes[1, 2]
    ax6.axis("off")
    unified_handles = [
        Line2D([0], [0], marker="o", color="w",
               markerfacecolor=OBS_COLOR, markeredgecolor=OBS_COLOR,
               markersize=8, label="Observations"),
        Line2D([0], [0], marker="D", color="w",
               markerfacecolor=REMOVED_PT_COLOR, markeredgecolor=REMOVED_PT_COLOR,
               markersize=8, label="Removed points"),
        Line2D([0], [0], color=FIT_LINE_COLOR, linewidth=1.8,
               label="OLS fit"),
        Patch(facecolor=RMSE_BAND_COLOR, alpha=0.35, edgecolor=RMSE_BAND_COLOR,
              label="\u00b1RMSE band"),
    ]
    ax6.legend(handles=unified_handles, fontsize=13, frameon=True,
               facecolor="white", edgecolor="#cccccc", framealpha=1.0,
               loc="center", handletextpad=0.6,
               borderpad=0.9, labelspacing=0.6,
               title="Legend", title_fontsize=14)

    plt.tight_layout(h_pad=3.5, w_pad=2.5)
    save_figure(fig, "Figure_pooled_scaling_all")
    plt.close(fig)


def figure_pooled_scaling_all_no_outliers():
    """Exact same figure as figure_pooled_scaling_all() -- same bands, same
    blue observation dots, same fit lines and equation boxes -- EXCEPT the
    red removed/influential points are not drawn. This is a separate figure;
    figure_pooled_scaling_all() itself is unchanged in its own output file."""
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch

    print("Figure — pooled power-law scaling, all equipment types (outliers hidden)")

    equipment_order = [
        ("Agitated Tanks",   "V", "-3"),
        ("Mixers",           "V", "-3"),
        ("Filtration Units", "A", "-2"),
        ("Dryers",           "A", "-2"),
        ("Calciners",        "V", "-3"),
    ]
    letters = ["a", "b", "c", "d", "e"]

    fig, axes = plt.subplots(2, 3, figsize=(14, 9))

    for pidx, (eq_name, kind, sup) in enumerate(equipment_order):
        row, col = divmod(pidx, 3)
        ax = axes[row, col]
        _draw_pooled_scaling_panel(ax, eq_name, kind, sup, letters[pidx],
                                   show_outliers=False)

    # ── Unified legend in cell (1, 2) -- no "Removed points" entry ─────────
    ax6 = axes[1, 2]
    ax6.axis("off")
    unified_handles = [
        Line2D([0], [0], marker="o", color="w",
               markerfacecolor=OBS_COLOR, markeredgecolor=OBS_COLOR,
               markersize=8, label="Observations"),
        Line2D([0], [0], color=FIT_LINE_COLOR, linewidth=1.8,
               label="OLS fit"),
        Patch(facecolor=RMSE_BAND_COLOR, alpha=0.35, edgecolor=RMSE_BAND_COLOR,
              label="\u00b1RMSE band"),
    ]
    ax6.legend(handles=unified_handles, fontsize=13, frameon=True,
               facecolor="white", edgecolor="#cccccc", framealpha=1.0,
               loc="center", handletextpad=0.6,
               borderpad=0.9, labelspacing=0.6,
               title="Legend", title_fontsize=14)

    plt.tight_layout(h_pad=3.5, w_pad=2.5)
    save_figure(fig, "Figure_pooled_scaling_all_no_outliers")
    plt.close(fig)


# ============================================================================= #
# FIGURE(S): SUBTYPE-LEVEL POWER-LAW SCALING — ONE FIGURE PER EQUIPMENT GROUP
#   Same style as figure_pooled_scaling_all but split by subtype.
#   Calciners: plain blue scatter only (no fit, no removal).
#   Layout: 3-column grid; legend in first empty cell, or as fig.legend if full.
# ============================================================================= #
def figure_subtype_scaling_by_group():
    """One figure per equipment group; one panel per subtype after outlier removal."""
    import statsmodels.api as sm
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch
    import pandas as pd
    import math

    print("Figure(s) — subtype-level scaling, one figure per equipment group")

    SKIP_GROUPS    = {"Calciners", "Agitated Tanks"}  # no subtypes
    NCOLS = 3

    for eq_name, (df, kind, sup) in EQUIPMENT.items():
        if eq_name in SKIP_GROUPS:
            continue
        subtypes = [t for t in df["Type"].unique()
                    if len(df[df["Type"] == t]) >= 4]
        n = len(subtypes)
        if n == 0:
            continue

        is_calciner = False  # Calciners skipped above
        pv_str  = "P/V" if kind == "V" else "P/A"
        cap_str = "V"   if kind == "V" else "A"
        exp_cap = "3" if kind == "V" else "2"
        x_unit  = f"m$^{{{exp_cap}}}$"
        y_unit  = f"kW$\\cdot$m$^{{{sup}}}$"

        ncols      = min(n, NCOLS)
        nrows      = math.ceil(n / ncols)
        total_cells = nrows * ncols
        has_empty  = (total_cells > n) and not is_calciner

        # Single-row figures get extra height so legend clears x-axis labels
        fig_h = 6.5 if nrows == 1 else 5 * nrows
        fig, axes = plt.subplots(nrows, ncols,
                                 figsize=(14, fig_h),
                                 squeeze=False)
        axes_flat = axes.flatten()
        letters   = [chr(ord("a") + i) for i in range(n)]

        # ── Legend proxy handles (built once, reused) ────────────────────────
        unified_handles = [
            Line2D([0], [0], marker="o", color="w",
                   markerfacecolor=OBS_COLOR, markeredgecolor=OBS_COLOR,
                   markersize=8, label="Observations"),
            Line2D([0], [0], marker="D", color="w",
                   markerfacecolor=REMOVED_PT_COLOR,
                   markeredgecolor=REMOVED_PT_COLOR,
                   markersize=8, label="Removed points"),
            Line2D([0], [0], color=FIT_LINE_COLOR,
                   linewidth=1.8, label="OLS fit"),
            Patch(facecolor=RMSE_BAND_COLOR, alpha=0.35,
                  edgecolor=RMSE_BAND_COLOR, label="\u00b1RMSE band"),
        ]

        for pidx, subtype in enumerate(subtypes):
            ax  = axes_flat[pidx]
            sub = df[df["Type"] == subtype].reset_index(drop=True)

            # Short display name
            short = (subtype.replace(" Mixers",  "").replace(" Filters", "")
                            .replace(" Dryers",  "").replace(" Calciners","")
                            .replace(" Units",   ""))

            if is_calciner:
                # ── Plain blue scatter — no removal, no fit ──────────────────
                ax.scatter(sub["logCap"].values, sub["logPV"].values,
                           color=OBS_COLOR, alpha=0.70, s=14,
                           linewidths=0.3, edgecolors=OBS_COLOR, zorder=2)
            else:
                # ── Identify & remove influentials ───────────────────────────
                inf_mask  = _identify_influentials(sub)
                sub_clean = sub.loc[~inf_mask].copy().reset_index(drop=True)
                if len(sub_clean) < 4:
                    sub_clean = sub.copy()
                    inf_mask  = pd.Series([False] * len(sub), index=sub.index)

                x_c = sub_clean["logCap"].values
                y_c = sub_clean["logPV"].values
                X   = np.vander(x_c, 2, increasing=True)
                mdl = sm.OLS(y_c, X).fit(cov_type="HC3")
                a, b   = float(mdl.params[0]), float(mdl.params[1])
                rmse   = float(np.sqrt(np.mean(mdl.resid**2)))
                r2_adj = float(mdl.rsquared_adj)

                # Observations
                ax.scatter(x_c, y_c, color=OBS_COLOR, alpha=0.70, s=14,
                           linewidths=0.3, edgecolors=OBS_COLOR, zorder=2)
                # Removed points
                removed = sub.loc[inf_mask]
                if len(removed) > 0:
                    ax.scatter(removed["logCap"].values, removed["logPV"].values,
                               color=REMOVED_PT_COLOR, alpha=0.95, s=26,
                               marker="D", linewidths=0.4,
                               edgecolors=REMOVED_PT_COLOR, zorder=4)
                # RMSE band + fit line
                xg = np.linspace(x_c.min() - 0.15, x_c.max() + 0.15, 300)
                yg = a + b * xg
                ax.fill_between(xg, yg - rmse, yg + rmse,
                                color=RMSE_BAND_COLOR, alpha=0.22, zorder=1)
                ax.plot(xg, yg, color=FIT_LINE_COLOR, linewidth=1.5, zorder=3)
                # ±RMSE annotation — top-right
                ax.text(0.97, 0.97, f"\u00b1RMSE = {rmse:.4f}",
                        transform=ax.transAxes, fontsize=11,
                        va="top", ha="right", color=RMSE_BAND_COLOR,
                        bbox=dict(facecolor="white", edgecolor=RMSE_BAND_COLOR,
                                  alpha=0.92, boxstyle="round,pad=0.35",
                                  linewidth=0.8))
                # Equation + R² box — lower-left
                sign    = "+" if b >= 0 else "\u2212"
                eq_line = (f"log({pv_str}) = {a:.4f} {sign} {abs(b):.4f}"
                           f" log({cap_str})\n"
                           f"$R^2_{{adj}}$ = {r2_adj:.4f}")
                ax.text(0.03, 0.04, eq_line,
                        transform=ax.transAxes, fontsize=11,
                        va="bottom", ha="left",
                        bbox=dict(facecolor="white", edgecolor="#aaaaaa",
                                  alpha=0.92, boxstyle="round,pad=0.4",
                                  linewidth=0.8))

            # Subtype name as panel title
            ax.set_title(short, fontsize=11, fontweight="bold", pad=5)
            ax.set_xlabel(f"log({cap_str})",
                          style="italic", fontsize=12)
            ax.set_ylabel(f"log({pv_str})",
                          style="italic", fontsize=12)
            ax.tick_params(axis="both", labelsize=11)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.annotate(f"({letters[pidx]})", xy=(-0.14, 1.04),
                        xycoords="axes fraction", ha="left", va="bottom",
                        fontsize=14, fontweight="bold", color="#333333")

        # ── Legend placement ─────────────────────────────────────────────────
        if has_empty:
            # Use first empty cell
            ax_leg = axes_flat[n]
            ax_leg.set_visible(True)
            ax_leg.axis("off")
            ax_leg.legend(handles=unified_handles, fontsize=11, frameon=True,
                          facecolor="white", edgecolor="#cccccc", framealpha=1.0,
                          loc="center", handletextpad=0.6,
                          borderpad=0.9, labelspacing=0.6,
                          title="Legend", title_fontsize=12)
            # Hide remaining empty cells
            for hidx in range(n + 1, total_cells):
                axes_flat[hidx].set_visible(False)
        elif not is_calciner:
            # No empty cell — horizontal legend inside the bottom margin.
            # bbox_to_anchor uses figure-fraction coords (0=bottom, 1=top),
            # so legend_y must be a small positive number within the margin.
            bottom_margin = 0.28 if nrows == 1 else 0.13
            legend_y      = 0.08 if nrows == 1 else 0.03
            fig.legend(handles=unified_handles, fontsize=11, frameon=True,
                       facecolor="white", edgecolor="#cccccc", framealpha=1.0,
                       loc="lower center", ncol=4,
                       handletextpad=0.6, borderpad=0.9, labelspacing=0.6,
                       bbox_to_anchor=(0.5, legend_y))
        else:
            # Calciner: hide any empty cells
            for hidx in range(n, total_cells):
                axes_flat[hidx].set_visible(False)

        safe = eq_name.replace(" ", "_")
        plt.tight_layout(h_pad=3.5, w_pad=2.5)
        if not has_empty and not is_calciner:
            plt.subplots_adjust(bottom=bottom_margin)
        save_figure(fig, f"Figure_subtype_scaling_{safe}")
        plt.close(fig)
        print(f"  saved: {eq_name} ({n} subtypes)")


# ============================================================================= #
# FIGURE 43  —  CAPACITY SIZE DISTRIBUTION (VENDOR DATASET COVERAGE)
#   5 panels: Agitated Tanks, Mixers, Filtration Units, Dryers, Calciners
#   Layout: 2 rows x 3 cols (5 panels + 1 legend cell)
#   Data: post-outlier-removal (using _identify_influentials per equipment)
#   X-axis: working capacity in raw units (m³ or m²), log-spaced bins
#   Y-axis: count
#   Goal: show vendor listings concentrated in mid-range commercial sizes,
#         with tails thinly sampled.
# ============================================================================= #
def figure_capacity_distribution():
    print("Figure 43 — capacity size distribution (vendor dataset coverage)")

    order   = ["Agitated Tanks", "Mixers", "Filtration Units", "Dryers", "Calciners"]
    letters = ["a", "b", "c", "d", "e"]

    cap_units = {
        "Agitated Tanks":   ("m$^3$", "Volume"),
        "Mixers":           ("m$^3$", "Volume"),
        "Filtration Units": ("m$^2$", "Area"),
        "Dryers":           ("m$^2$", "Area"),
        "Calciners":        ("m$^3$", "Volume"),
    }

    def _fmt_cap(v):
        if v < 0.01:   return f"{v:.3f}"
        elif v < 0.1:  return f"{v:.2f}"
        elif v < 10:   return f"{v:.1f}"
        elif v < 100:  return f"{v:.0f}"
        else:          return f"{v:.0f}"

    def _build_clean(name, df):
        """Apply per-subtype outlier removal matching each notebook's method."""
        subtypes = df["Type"].unique()
        keep_idx = []
        for t in subtypes:
            sub = df[df["Type"] == t].copy()
            if len(sub) < 4:
                keep_idx.extend(sub.index.tolist())
                continue
            infl = _identify_influentials(sub, equipment_name=name)
            keep_idx.extend(sub.index[~infl].tolist())
        return df.loc[keep_idx].reset_index(drop=True)

    def _draw_capacity_hist(ax, cap_clean, letter, cap_unit, cap_word, display_name):
        """Draw one capacity histogram panel."""
        log_cap = np.log10(cap_clean)

        # Freedman-Diaconis bins on log scale
        n_bins = fd_bins(log_cap)
        bin_edges       = np.linspace(log_cap.min(), log_cap.max(), n_bins + 1)
        counts, _       = np.histogram(log_cap, bins=bin_edges)
        bin_centers_log = 0.5 * (bin_edges[:-1] + bin_edges[1:])
        bin_centers_raw = 10 ** bin_centers_log
        bin_widths      = np.diff(bin_edges)

        # Bars
        ax.bar(
            bin_centers_log, counts,
            width=bin_widths,
            color=OBS_COLOR, alpha=0.55,
            edgecolor=OBS_COLOR, linewidth=0.45,
            align="center",
        )

        # Count labels on top of bars
        for cx, cnt in zip(bin_centers_log, counts):
            if cnt > 0:
                ax.text(cx, cnt + 0.15, str(cnt),
                        ha="center", va="bottom",
                        fontsize=11, color="black")

        # 5th / 95th percentile lines
        p5  = np.percentile(log_cap, 5)
        p95 = np.percentile(log_cap, 95)
        ax.axvline(p5,  color=FIT_LINE_COLOR, linewidth=1.0,
                   linestyle="--", alpha=0.70, zorder=3)
        ax.axvline(p95, color=FIT_LINE_COLOR, linewidth=1.0,
                   linestyle=":",  alpha=0.70, zorder=3)

        # X-ticks in raw units
        tick_labels = [_fmt_cap(v) for v in bin_centers_raw]
        ax.set_xticks(bin_centers_log)
        ax.set_xticklabels(tick_labels, rotation=45, ha="right", fontsize=14)
        ax.tick_params(axis="y", labelsize=16)

        # Summary stats below x-axis label
        n_total = len(cap_clean)
        c_min   = cap_clean.min()
        c_max   = cap_clean.max()
        c_med   = np.median(cap_clean)
        ax.set_xlabel(
            f"{cap_word} [{cap_unit}]\n"
            f"$n$ = {n_total},  "
            f"range: {_fmt_cap(c_min)}–{_fmt_cap(c_max)} {cap_unit},  "
            f"median: {_fmt_cap(c_med)} {cap_unit}",
            style="italic", fontsize=14,
        )
        ax.set_ylabel("Count", style="italic", fontsize=16)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

        # Panel letter — outside top-left
        ax.annotate(
            f"({letter}) {display_name}",
            xy=(-0.02, 1.05), xycoords="axes fraction",
            ha="left", va="bottom",
            fontsize=14, fontweight="bold", color="#333333",
        )

    # ── Layout: 2 rows × 3 cols ───────────────────────────────────────────────
    fig, axes = plt.subplots(2, 3, figsize=(16.5, 8.5))
    axes_flat = axes.flatten()

    panel_map = {
        "Agitated Tanks":   0,
        "Mixers":           1,
        "Filtration Units": 2,
        "Dryers":           3,
        "Calciners":        4,   # bottom-left; bottom-right used for legend
    }

    for name, letter in zip(order, letters):
        df_raw, kind, sup = EQUIPMENT[name]
        cap_unit, cap_word = cap_units[name]

        df_clean  = _build_clean(name, df_raw)
        cap_clean = df_clean["Capacity"].values

        ax   = axes_flat[panel_map[name]]
        _draw_capacity_hist(ax, cap_clean, letter, cap_unit, cap_word,
                             DISPLAY_NAME.get(name, name))


    # ── Bottom-right cell: legend for percentile lines ────────────────────────
    from matplotlib.lines import Line2D
    ax_leg = axes_flat[5]
    ax_leg.set_visible(True)
    ax_leg.axis("off")
    legend_handles = [
        Line2D([0], [0], color=FIT_LINE_COLOR, linewidth=1.2,
               linestyle="--", label="5th percentile"),
        Line2D([0], [0], color=FIT_LINE_COLOR, linewidth=1.2,
               linestyle=":",  label="95th percentile"),
    ]
    ax_leg.legend(
        handles=legend_handles,
        fontsize=12, frameon=True,
        facecolor="white", edgecolor="#cccccc",
        loc="center",
        title="Capacity range markers",
        title_fontsize=12,
    )

    plt.tight_layout(h_pad=3.5, w_pad=2.5)
    plt.subplots_adjust(left=0.045, bottom=0.14)
    save_figure(fig, "Figure_capacity_distribution")
    plt.close(fig)


if __name__ == "__main__":
    figure_s1_distributions()
    figure_s2_power_capacity_trends()
    figure_influence_diagnostics()
    figure_scaling_before_after()
    figure_pooled_capacity_boxplots()
    figure_subtype_capacity_boxplots()
    figure_assumption_checks()
    figure_multi_subtype_scaling()
    figure_pooled_scaling_all()
    figure_pooled_scaling_all_no_outliers()
    figure_subtype_scaling_by_group()
    figure_capacity_distribution()
    # --- add further figure_*() calls here as we build them ---
    print("Done.")
