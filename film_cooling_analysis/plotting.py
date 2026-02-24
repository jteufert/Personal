"""
Visualization for Film Cooling Analysis Results

Generates a multi-panel figure summarizing the thermal analysis:
  1. Nozzle contour and Mach number
  2. Heat transfer coefficient h_g(x)
  3. Temperatures: T_aw, T_aw_fc, T_film along the nozzle
  4. Film cooling effectiveness eta(x)
  5. Wall temperature: with film+rad vs. rad-only cooling
  6. Heat flux comparison
"""

import numpy as np
import matplotlib
matplotlib.use("Agg")   # Non-interactive backend for headless environments
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from matplotlib.patches import Patch
from typing import Optional


def _throat_line(ax, label: bool = True):
    """Draw a vertical dashed line at the throat (x=0)."""
    kw = dict(color="gray", lw=1.0, ls="--", alpha=0.7)
    ax.axvline(0.0, **kw)
    if label:
        ax.text(0.0, ax.get_ylim()[1] * 0.95, "throat",
                ha="center", va="top", fontsize=7, color="gray")


def _x_mm(x: np.ndarray) -> np.ndarray:
    """Convert axial coordinate from metres to millimetres."""
    return x * 1e3


def plot_results(
    results: dict,
    nozzle,
    config: dict,
    save_path: Optional[str] = None,
    show: bool = False,
) -> plt.Figure:
    """
    Generate the main results figure.

    Parameters
    ----------
    results : dict
        Output from ThrusterAnalysis.run()
    nozzle : NozzleGeometry
        Nozzle geometry object
    config : dict
        Analysis configuration dict
    save_path : str or None
        If provided, save the figure to this path (PNG recommended)
    show : bool
        If True, call plt.show() (requires a display)

    Returns
    -------
    matplotlib.figure.Figure
    """
    x    = results["x"]
    xmm  = _x_mm(x)

    fig = plt.figure(figsize=(14, 16))
    fig.suptitle(
        f"Film Cooling Analysis – {config.get('ox_name','GOX')} / "
        f"{config.get('fuel_name','Ethanol')}\n"
        f"Pc = {config['Pc']:.1f} bar | O/F = {config['MR']:.2f} | "
        f"ε = {config['eps']:.0f} | "
        f"Film = {config['film_fraction']*100:.0f}% fuel | "
        f"F = {config['thrust']:.0f} N",
        fontsize=11, y=0.98,
    )

    gs = gridspec.GridSpec(3, 2, figure=fig, hspace=0.45, wspace=0.35)

    # ------------------------------------------------------------------
    # Panel 1: Nozzle contour + Mach number
    # ------------------------------------------------------------------
    ax1a = fig.add_subplot(gs[0, 0])
    r_mm = nozzle.radius(x) * 1e3
    ax1a.fill_between(xmm,  r_mm,  r_mm * 0 + r_mm.max() * 1.2,
                      alpha=0.15, color="steelblue", label="Wall")
    ax1a.fill_between(xmm, -r_mm, -r_mm * 0 - r_mm.max() * 1.2,
                      alpha=0.15, color="steelblue")
    ax1a.plot(xmm,  r_mm, "b-", lw=1.5)
    ax1a.plot(xmm, -r_mm, "b-", lw=1.5)
    ax1a.axvline(0, color="gray", ls="--", lw=1, alpha=0.7)
    ax1a.set_xlim(xmm[0], xmm[-1])
    ax1a.set_xlabel("Axial position [mm]")
    ax1a.set_ylabel("Radius [mm]")
    ax1a.set_title("Nozzle Contour")
    ax1a.set_aspect("equal", adjustable="datalim")
    ax1a.text(0, 0, "throat", ha="center", va="center",
              fontsize=7, color="gray")

    # Mach number on twin axis
    ax1b = ax1a.twinx()
    ax1b.plot(xmm, results["M"], "r-", lw=1.5, label="Mach")
    ax1b.set_ylabel("Mach number [-]", color="red")
    ax1b.tick_params(axis="y", labelcolor="red")
    ax1b.set_ylim(0, results["M"].max() * 1.1)

    # ------------------------------------------------------------------
    # Panel 2: Heat transfer coefficient
    # ------------------------------------------------------------------
    ax2 = fig.add_subplot(gs[0, 1])
    ax2.plot(xmm, results["h_g"] / 1e3, "C1-", lw=1.8)
    ax2.set_xlabel("Axial position [mm]")
    ax2.set_ylabel("h_g [kW/(m²·K)]")
    ax2.set_title("Convective Heat Transfer Coefficient (Bartz)")
    ax2.axvline(0, color="gray", ls="--", lw=1, alpha=0.7)
    ax2.text(0, ax2.get_ylim()[1] * 0.05, " throat", fontsize=7, color="gray")
    ax2.set_xlim(xmm[0], xmm[-1])
    ax2.grid(True, alpha=0.3)

    # ------------------------------------------------------------------
    # Panel 3: Temperature profiles
    # ------------------------------------------------------------------
    ax3 = fig.add_subplot(gs[1, 0])
    ax3.plot(xmm, results["T_aw"],    "r-",   lw=1.5, label="$T_{aw}$ (no film)")
    ax3.plot(xmm, results["T_aw_fc"], "C0-",  lw=1.5, label="$T_{aw,fc}$ (with film)")
    ax3.plot(xmm, results["T_film"],  "g--",  lw=1.2, label="$T_{film}$")
    ax3.axhline(config["T_film_in"], color="green", ls=":", lw=1,
                label=f"Film inlet ({config['T_film_in']:.0f} K)")
    ax3.axvline(-nozzle.L_conv * 1e3, color="purple", ls=":", lw=1.2,
                label="Film injection")
    ax3.axvline(0, color="gray", ls="--", lw=1, alpha=0.7)
    ax3.set_xlabel("Axial position [mm]")
    ax3.set_ylabel("Temperature [K]")
    ax3.set_title("Gas-Side Temperatures")
    ax3.legend(fontsize=7, loc="upper right")
    ax3.set_xlim(xmm[0], xmm[-1])
    ax3.grid(True, alpha=0.3)

    # ------------------------------------------------------------------
    # Panel 4: Film effectiveness
    # ------------------------------------------------------------------
    ax4 = fig.add_subplot(gs[1, 1])
    ax4.plot(xmm, results["eta"] * 100, "C2-", lw=1.8)
    ax4.fill_between(xmm, results["eta"] * 100, alpha=0.2, color="green")
    ax4.axvline(-nozzle.L_conv * 1e3, color="purple", ls=":", lw=1.2,
                label="Film injection")
    ax4.axvline(0, color="gray", ls="--", lw=1, alpha=0.7)
    ax4.set_xlabel("Axial position [mm]")
    ax4.set_ylabel("Effectiveness η [%]")
    ax4.set_title("Film Cooling Effectiveness")
    ax4.set_ylim(0, 105)
    ax4.legend(fontsize=7)
    ax4.set_xlim(xmm[0], xmm[-1])
    ax4.grid(True, alpha=0.3)

    # Annotate film fraction burnout
    eta_thresh = 5.0  # 5% effectiveness threshold
    burnout_mask = (results["eta"] * 100 < eta_thresh) & (x > -nozzle.L_conv)
    if burnout_mask.any():
        x_burnout = xmm[burnout_mask][0]
        ax4.axvline(x_burnout, color="red", ls="-.", lw=1.0, alpha=0.7)
        ax4.text(x_burnout, 50, f" η<{eta_thresh:.0f}%\n x={x_burnout:.1f}mm",
                 fontsize=7, color="red")

    # ------------------------------------------------------------------
    # Panel 5: Wall temperature comparison
    # ------------------------------------------------------------------
    ax5 = fig.add_subplot(gs[2, 0])
    ax5.plot(xmm, results["T_wall_fc"],   "C0-",  lw=1.8,
             label="Film + Radiation cooling")
    ax5.plot(xmm, results["T_wall_nofc"], "r--",  lw=1.5,
             label="Radiation only (no film)")
    ax5.axhline(config["T_wall_max"], color="k", ls=":", lw=1.5,
                label=f"Material limit ({config['T_wall_max']:.0f} K)")
    ax5.axvline(-nozzle.L_conv * 1e3, color="purple", ls=":", lw=1.2)
    ax5.axvline(0, color="gray", ls="--", lw=1, alpha=0.7)
    ax5.set_xlabel("Axial position [mm]")
    ax5.set_ylabel("Wall temperature [K]")
    ax5.set_title("Wall Temperature Distribution")
    ax5.legend(fontsize=7)
    ax5.set_xlim(xmm[0], xmm[-1])
    ax5.grid(True, alpha=0.3)

    # Shade the region above material limit
    T_lim = config["T_wall_max"]
    ax5.fill_between(xmm, T_lim, results["T_wall_fc"],
                     where=results["T_wall_fc"] > T_lim,
                     color="red", alpha=0.3, label="Exceeds limit")

    # ------------------------------------------------------------------
    # Panel 6: Heat flux comparison
    # ------------------------------------------------------------------
    ax6 = fig.add_subplot(gs[2, 1])
    ax6.plot(xmm, results["q_conv_fc"]   / 1e6, "C0-",  lw=1.8,
             label="With film cooling")
    ax6.plot(xmm, results["q_conv_nofc"] / 1e6, "r--",  lw=1.5,
             label="No film cooling")
    ax6.axvline(-nozzle.L_conv * 1e3, color="purple", ls=":", lw=1.2,
                label="Film injection")
    ax6.axvline(0, color="gray", ls="--", lw=1, alpha=0.7)
    ax6.set_xlabel("Axial position [mm]")
    ax6.set_ylabel("Heat flux [MW/m²]")
    ax6.set_title("Wall Heat Flux")
    ax6.legend(fontsize=7)
    ax6.set_xlim(xmm[0], xmm[-1])
    ax6.grid(True, alpha=0.3)

    # ------------------------------------------------------------------
    # Footer annotation
    # ------------------------------------------------------------------
    fig.text(
        0.5, 0.01,
        f"Throat: D = {results['D_throat']*1e3:.2f} mm | "
        f"ṁ_total = {results['m_dot_total']*1e3:.2f} g/s | "
        f"ṁ_film = {results['m_dot_film']*1e3:.2f} g/s | "
        f"ε_wall = {config['emissivity']:.2f} | "
        f"T_sink = {config['T_background']:.0f} K",
        ha="center", va="bottom", fontsize=8, style="italic",
    )

    if save_path:
        fig.savefig(save_path, dpi=150, bbox_inches="tight")
        print(f"  Plot saved to: {save_path}")

    if show:
        plt.show()

    return fig


def plot_nozzle_contour(nozzle, save_path: Optional[str] = None) -> plt.Figure:
    """
    Simple standalone nozzle contour plot with area ratio overlay.
    """
    x = nozzle.x
    xmm = x * 1e3
    r_mm = nozzle.radius(x) * 1e3
    AR = nozzle.area_ratio(x)

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 6))

    ax1.fill_between(xmm,  r_mm, color="steelblue", alpha=0.3)
    ax1.fill_between(xmm, -r_mm, color="steelblue", alpha=0.3)
    ax1.plot(xmm,  r_mm, "b-", lw=2)
    ax1.plot(xmm, -r_mm, "b-", lw=2)
    ax1.axvline(0, color="gray", ls="--", alpha=0.7)
    ax1.set_ylabel("Radius [mm]")
    ax1.set_title("Nozzle Contour")
    ax1.set_aspect("equal", adjustable="datalim")
    ax1.grid(True, alpha=0.3)

    ax2.semilogy(xmm, AR, "C1-", lw=2)
    ax2.axvline(0, color="gray", ls="--", alpha=0.7)
    ax2.set_xlabel("Axial position [mm]")
    ax2.set_ylabel("Area ratio A/A*")
    ax2.set_title("Area Ratio Profile")
    ax2.grid(True, alpha=0.3)

    plt.tight_layout()
    if save_path:
        fig.savefig(save_path, dpi=150, bbox_inches="tight")
    return fig
