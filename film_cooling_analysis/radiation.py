"""
Radiative Heat Transfer for Film-Cooled Thruster

Models external thermal radiation from the thruster wall to the
surrounding environment (space or laboratory).

The net radiative heat flux leaving the outer wall surface is:

  q_rad = epsilon * sigma_SB * (T_wall^4 - T_background^4)

where:
  epsilon    : wall surface emissivity [-]
  sigma_SB   : Stefan-Boltzmann constant = 5.6704e-8 W/(m²·K⁴)
  T_wall     : outer wall temperature [K]
  T_background: sink temperature [K] (4 K for deep space, ~300 K for lab)

For a vacuum environment radiating to deep space, T_background ≈ 4 K,
and the T_bg^4 term is negligible.
"""

import numpy as np

SIGMA_SB = 5.6704e-8   # Stefan-Boltzmann constant [W/(m²·K⁴)]


def radiation_flux(
    T_wall: float,
    emissivity: float,
    T_background: float = 4.0,
) -> float:
    """
    Net radiative heat flux leaving the outer wall [W/m²].

    q_rad = epsilon * sigma_SB * (T_wall^4 - T_background^4)

    Parameters
    ----------
    T_wall : float
        Wall outer surface temperature [K]
    emissivity : float
        Surface emissivity (0 = perfect mirror, 1 = blackbody)
    T_background : float
        Radiation sink temperature [K]

    Returns
    -------
    float
        Net radiative flux from the wall [W/m²] (positive = heat out)
    """
    return emissivity * SIGMA_SB * (T_wall ** 4 - T_background ** 4)


def radiation_flux_array(
    T_wall: np.ndarray,
    emissivity: float,
    T_background: float = 4.0,
) -> np.ndarray:
    """Vectorized radiation_flux over an array of wall temperatures."""
    T_wall = np.asarray(T_wall, dtype=float)
    return emissivity * SIGMA_SB * (T_wall ** 4 - T_background ** 4)


def max_radiation_temperature(
    heat_flux: float,
    emissivity: float,
    T_background: float = 4.0,
) -> float:
    """
    Maximum equilibrium wall temperature for a given incoming heat flux
    if the only cooling is radiation.

    Solves: q = epsilon * sigma * (T^4 - T_bg^4)
    => T = (q / (epsilon * sigma) + T_bg^4)^(1/4)

    Parameters
    ----------
    heat_flux : float
        Incoming heat flux to be rejected [W/m²]
    emissivity : float
        Surface emissivity [-]
    T_background : float
        Sink temperature [K]

    Returns
    -------
    float
        Equilibrium wall temperature [K]
    """
    T4 = heat_flux / (emissivity * SIGMA_SB) + T_background ** 4
    return T4 ** 0.25
