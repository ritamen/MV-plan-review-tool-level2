"""
regression_verifier.py
----------------------
Pure Python/NumPy/SciPy OLS regression engine for M&V plan verification.

For each EEM that uses regression-based routine adjustment:
  1. Runs OLS regression independently (baseline kWh ~ independent variable)
  2. Computes R², CV(RMSE), NMBE, t-statistic, p-value, model standard error
  3. Compares computed stats against values reported in the M&V plan (±2% tolerance)
  4. Assesses against IPMVP Core Concepts + ASHRAE Guideline 14 thresholds

IPMVP / ASHRAE Guideline 14 thresholds (monthly data):
  R²          > 0.75
  |t-stat|    > 2.0   (for every independent variable)
  CV(RMSE)    ≤ 20%
  NMBE        ≤ ±5%
  Expected savings > 2 × model standard error
"""

try:
    import numpy as np
    from scipy import stats as scipy_stats
    _SCIPY_AVAILABLE = True
except ImportError:
    _SCIPY_AVAILABLE = False

# ── IPMVP / ASHRAE Guideline 14 thresholds (monthly data) ────────────────────
R2_MIN       = 0.75
T_STAT_MIN   = 2.0
CVRMSE_MAX   = 20.0   # percent
NMBE_MAX_ABS = 5.0    # percent (absolute value)
MATCH_TOL    = 0.02   # 2% relative tolerance for computed vs reported comparison


def _check_match(computed, reported, tol=MATCH_TOL):
    """Return True/False/None (N/A) for computed vs reported comparison."""
    if computed is None or reported is None:
        return None
    if reported == 0:
        return abs(computed) < 1e-9
    return abs(computed - reported) / abs(reported) <= tol


def _run_ols(baseline_kwh, indep_values):
    """
    Run OLS regression: baseline_kwh ~ indep_values.

    Returns a dict of computed statistics.
    Raises ValueError on invalid input.
    """
    if not _SCIPY_AVAILABLE:
        raise ImportError("numpy and scipy are required for regression verification.")

    x = np.array(indep_values, dtype=float)
    y = np.array(baseline_kwh, dtype=float)

    if len(x) != len(y):
        raise ValueError(
            f"Array length mismatch: baseline has {len(y)} values, "
            f"independent variable has {len(x)} values."
        )
    if len(y) < 3:
        raise ValueError(
            f"Insufficient data points: need ≥ 3, got {len(y)}."
        )

    mean_y = float(np.mean(y))
    if mean_y == 0:
        raise ValueError("Mean of baseline_kwh is zero — CV(RMSE) and NMBE undefined.")

    # SciPy simple linear regression
    slope, intercept, r_value, p_value, slope_stderr = scipy_stats.linregress(x, y)

    r_squared = float(r_value ** 2)
    y_pred    = slope * x + intercept
    residuals = y - y_pred
    n         = len(y)

    # Model standard error (RMSE with n-2 degrees of freedom)
    sse           = float(np.sum(residuals ** 2))
    model_std_err = float(np.sqrt(sse / (n - 2)))

    # CV(RMSE) = model_std_err / mean(y)  ×  100  [%]
    cv_rmse = (model_std_err / mean_y) * 100.0

    # NMBE = Σ(predicted - actual) / (n × mean_y)  ×  100  [%]
    # For a fitted OLS model this is near zero by construction;
    # computed for completeness and cross-check against any reported value.
    nmbe = (float(np.sum(y_pred - y)) / (n * mean_y)) * 100.0

    # t-statistic for the slope (independent variable)
    t_stat = float(slope / slope_stderr) if slope_stderr != 0 else 0.0

    # t-statistic for the intercept
    # Use OLS normal equations: Var(b) = MSE * (X'X)^-1, intercept is index [0]
    try:
        X = np.column_stack([np.ones(n), x])
        XtX_inv = np.linalg.inv(X.T @ X)
        intercept_stderr  = float(np.sqrt(model_std_err ** 2 * XtX_inv[0, 0]))
        intercept_t_stat  = float(intercept / intercept_stderr) if intercept_stderr != 0 else 0.0
        intercept_p_value = float(2 * scipy_stats.t.sf(abs(intercept_t_stat), df=n - 2))
    except np.linalg.LinAlgError:
        intercept_stderr  = None
        intercept_t_stat  = None
        intercept_p_value = None

    return {
        "intercept":          float(intercept),
        "slope":              float(slope),
        "r_squared":          r_squared,
        "model_std_err":      model_std_err,
        "slope_stderr":       float(slope_stderr),
        "t_stat":             t_stat,
        "p_value":            float(p_value),
        "intercept_stderr":   intercept_stderr,
        "intercept_t_stat":   intercept_t_stat,
        "intercept_p_value":  intercept_p_value,
        "cv_rmse":            cv_rmse,
        "nmbe":               nmbe,
        "n":                  n,
        "mean_y":             mean_y,
    }


def verify_eem(
    eem_name: str,
    baseline_kwh: list,
    indep_values: list,
    reported_stats: dict = None,
    expected_savings_kwh: float = None,
) -> dict:
    """
    Full regression verification for one EEM.

    Parameters
    ----------
    eem_name             : str
    baseline_kwh         : list of float  — monthly baseline energy (kWh)
    indep_values         : list of float  — monthly independent variable (e.g. CDD)
    reported_stats       : dict           — stats from M&V plan, keys:
                             r_squared, cv_rmse, t_stat, p_value, model_std_err
                             (each float or None if not reported)
    expected_savings_kwh : float or None  — annual expected savings

    Returns
    -------
    dict with keys:
      eem_name, computed, comparison, thresholds, overall_pass,
      stats_mismatch, error
    """
    reported = reported_stats or {}

    # ── Run OLS ──────────────────────────────────────────────────────────────
    try:
        computed = _run_ols(baseline_kwh, indep_values)
    except Exception as exc:
        return {
            "eem_name":      eem_name,
            "computed":      None,
            "comparison":    {},
            "thresholds":    {},
            "overall_pass":  False,
            "stats_mismatch": [],
            "error":         str(exc),
        }

    # ── Compare computed vs reported ─────────────────────────────────────────
    comparison = {
        "R²": {
            "computed": computed["r_squared"],
            "reported": reported.get("r_squared"),
            "match":    _check_match(computed["r_squared"], reported.get("r_squared")),
        },
        "CV(RMSE) %": {
            "computed": computed["cv_rmse"],
            "reported": reported.get("cv_rmse"),
            "match":    _check_match(computed["cv_rmse"], reported.get("cv_rmse")),
        },
        "t-statistic": {
            "computed": computed["t_stat"],
            "reported": reported.get("t_stat"),
            "match":    _check_match(computed["t_stat"], reported.get("t_stat")),
        },
        "p-value": {
            "computed": computed["p_value"],
            "reported": reported.get("p_value"),
            "match":    _check_match(computed["p_value"], reported.get("p_value")),
        },
        "Model Std Error": {
            "computed": computed["model_std_err"],
            "reported": reported.get("model_std_err"),
            "match":    _check_match(computed["model_std_err"], reported.get("model_std_err")),
        },
    }

    stats_mismatch = [k for k, v in comparison.items() if v["match"] is False]

    # ── IPMVP / ASHRAE Guideline 14 threshold assessment ────────────────────
    r2                = computed["r_squared"]
    t_stat            = computed["t_stat"]
    cvrmse            = computed["cv_rmse"]
    nmbe              = computed["nmbe"]
    mse               = computed["model_std_err"]
    intercept_t_stat  = computed["intercept_t_stat"]

    thresholds = {
        "R² > 0.75": {
            "value":     r2,
            "threshold": f"> {R2_MIN}",
            "passes":    r2 > R2_MIN,
            "note":      "IPMVP minimum model fit",
        },
        "|t-stat| > 2": {
            "value":     t_stat,
            "threshold": f"|t| > {T_STAT_MIN}",
            "passes":    abs(t_stat) > T_STAT_MIN,
            "note":      "Statistical significance of independent variable (slope)",
        },
        "CV(RMSE) ≤ 20%": {
            "value":     cvrmse,
            "threshold": f"≤ {CVRMSE_MAX:.0f}% (monthly)",
            "passes":    cvrmse <= CVRMSE_MAX,
            "note":      "ASHRAE Guideline 14 monthly data",
        },
        "NMBE ≤ ±5%": {
            "value":     nmbe,
            "threshold": f"≤ ±{NMBE_MAX_ABS:.0f}% (monthly)",
            "passes":    abs(nmbe) <= NMBE_MAX_ABS,
            "note":      "ASHRAE Guideline 14 monthly data",
        },
    }

    if expected_savings_kwh is not None:
        threshold_val = 2.0 * mse
        thresholds["Savings > 2×StdErr"] = {
            "value":     expected_savings_kwh,
            "threshold": f"> 2 × {mse:,.0f} = {threshold_val:,.0f} kWh",
            "passes":    expected_savings_kwh > threshold_val,
            "note":      "IPMVP: savings must exceed twice the model standard error",
        }

    overall_pass = all(v["passes"] for v in thresholds.values())

    return {
        "eem_name":      eem_name,
        "computed":      computed,
        "comparison":    comparison,
        "thresholds":    thresholds,
        "overall_pass":  overall_pass,
        "stats_mismatch": stats_mismatch,
        "error":         None,
    }


def verify_all(eem_data: list) -> list:
    """
    Run verify_eem() for each EEM in eem_data.

    Parameters
    ----------
    eem_data : list of dicts, each with:
        eem_name             str
        baseline_kwh         list of float
        indep_values         list of float
        reported_stats       dict (optional)
        expected_savings_kwh float (optional)

    Returns list of result dicts from verify_eem().
    """
    if not eem_data:
        return []
    return [
        verify_eem(
            eem_name=d.get("eem_name", "Unknown EEM"),
            baseline_kwh=d.get("baseline_kwh", []),
            indep_values=d.get("indep_values", []),
            reported_stats=d.get("reported_stats"),
            expected_savings_kwh=d.get("expected_savings_kwh"),
        )
        for d in eem_data
    ]
