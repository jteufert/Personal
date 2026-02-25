"""
Film Cooling Analysis – GOX/Ethanol Thruster
=============================================

Entry point for the film cooling thermal analysis tool.

This script analyses a small GOX/Ethanol thruster typical of a university
CubeSat or small satellite propulsion system. It uses NASA CEA combustion
data (via rocketcea) to obtain gas properties, applies the Bartz (1957)
convective heat transfer correlation along the nozzle contour, integrates
the film coolant energy balance ODE, and solves for the steady-state wall
temperature accounting for both film and radiative cooling.

Usage
-----
    python main.py

The analysis prints a summary to stdout and saves a results figure to
film_cooling_results.png in the current directory.

Configuration
-------------
Edit the `config` dictionary below to change thruster parameters.
Key parameters to vary:
  - film_fraction: try 0.05-0.20 to see its effect on wall temperature
  - eps: higher area ratio pushes more heat flux to the throat region
  - Pc: higher pressure increases h_g and peak heat flux
  - emissivity: important material selection variable
"""

import sys
import os

# Allow running from the repo root without installing the package
sys.path.insert(0, os.path.dirname(__file__))

from film_cooling_analysis.cea_interface import CEAInterface
from film_cooling_analysis.analysis import ThrusterAnalysis
from film_cooling_analysis.plotting import plot_results

# ---------------------------------------------------------------------------
# Thruster Configuration
# ---------------------------------------------------------------------------

config = {
    # -- Propellants --------------------------------------------------------
    "ox_name":   "GOX",      # Gaseous oxygen
    "fuel_name": "Ethanol",  # Ethanol (C2H5OH)

    # -- Operating conditions -----------------------------------------------
    "Pc":  130 / 14.5038,  # Chamber pressure [bar]  (130 psi)
    "MR":   1.4,           # O/F mixture ratio (mass)
    #   Stoichiometric O/F for Ethanol ≈ 2.09; run fuel-rich for cooling

    # -- Nozzle sizing: set D_throat to fix geometry (thrust becomes output)
    #    or remove D_throat and set thrust [N] to size the throat instead.
    "D_throat": 0.004,     # Throat diameter [m]  (4 mm)

    # -- Nozzle geometry ----------------------------------------------------
    "eps":           50.0,   # Exit-to-throat area ratio Ae/At
    "Dc_Dt":          3.0,   # Chamber/throat diameter ratio
    "L_chamber":      0.05,  # Cylindrical chamber length [m]
    "alpha_conv_deg": 30.0,  # Convergent half-angle [deg]
    "alpha_div_deg":  15.0,  # Divergent half-angle [deg]  (conical nozzle)
    "rc_over_Rt":     0.8,   # Throat curvature / throat radius
    "n_points":      500,    # Axial analysis stations

    # -- Film cooling -------------------------------------------------------
    "film_fraction": 0.20,   # Film flow / total fuel flow (20%)
    #   Film coolant is ethanol vapour injected at the start of the convergent
    #   section along the inner wall.
    "Cp_film":      2570.0,  # Ethanol vapour Cp ≈ 2570 J/(kg·K) at ~400 K
    "T_film_in":     300.0,  # Film inlet temperature [K]

    # -- Wall properties ----------------------------------------------------
    "emissivity":  0.85,     # Oxidised stainless steel / Inconel outer surface
    "T_wall_max": 1644.0,    # Allowable wall temperature [K]  (1371 °C / 1500 °F)
    #   Inconel 625: ~1644 K (1500 °F) | Carbon-Carbon: ~2200 K | Rhenium: ~2700 K

    # -- Environment --------------------------------------------------------
    "T_background": 4.0,     # Deep-space radiation sink [K]
    #   Use ~300 K for a ground-test in a vacuum chamber
}

# ---------------------------------------------------------------------------
# Run analysis
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("\nFilm Cooling Analysis for GOX/Ethanol Thruster")
    print("=" * 50)

    # Print CEA combustion summary first
    cea = CEAInterface(config["ox_name"], config["fuel_name"])
    cea.print_summary(config["Pc"], config["MR"], config["eps"])

    # Run full thermal analysis
    print("Running thermal analysis...")
    analysis = ThrusterAnalysis(config)
    results  = analysis.run()

    analysis.nozzle.print_summary()
    analysis.print_summary()

    # Save results plot
    out_path = os.path.join(os.path.dirname(__file__), "film_cooling_results.png")
    fig = plot_results(
        results   = results,
        nozzle    = analysis.nozzle,
        config    = config,
        save_path = out_path,
        show      = False,
    )
    print(f"\nAnalysis complete. Results figure: {out_path}")
