"""
Hot-Gas-Side Heat Transfer: Bartz Equation

Implements the Bartz (1957) correlation for convective heat transfer
coefficients in rocket engine nozzles, along with adiabatic wall
temperature (recovery temperature) calculations.

References:
  Bartz, D. R. (1957). "A Simple Equation for Rapid Estimation of Rocket
    Nozzle Convective Heat Transfer Coefficients." Jet Propulsion, 27(1).
  Huzel, D. K. & Huang, D. H. (1992). Modern Engineering for Design of
    Liquid-Propellant Rocket Engines. AIAA.
"""

import numpy as np
from .nozzle import NozzleGeometry, compute_mach_profile, isentropic_temperature


def bartz_sigma(
    T_wall: float, T_chamber: float, M: float, gamma: float
) -> float:
    """
    Bartz correction factor sigma for property variation across the
    boundary layer between the wall and the core flow.

    sigma = [0.5*(T_w/T_c)*(1 + (gamma-1)/2*M^2) + 0.5]^(-0.68)
          * [1 + (gamma-1)/2*M^2]^(-0.12)

    Parameters
    ----------
    T_wall : float
        Wall temperature estimate [K] (iterate to converge)
    T_chamber : float
        Chamber (stagnation) temperature [K]
    M : float
        Local Mach number [-]
    gamma : float
        Specific heat ratio [-]

    Returns
    -------
    float
        Dimensionless correction factor sigma
    """
    factor = 1.0 + (gamma - 1.0) / 2.0 * M ** 2
    sigma = (
        (0.5 * (T_wall / T_chamber) * factor + 0.5) ** (-0.68)
        * factor ** (-0.12)
    )
    return float(sigma)


def bartz_coefficient(
    D_throat: float,
    r_c: float,
    Pc_Pa: float,
    c_star: float,
    mu_chamber: float,
    Cp_chamber: float,
    Pr_chamber: float,
    T_chamber: float,
    gamma: float,
    M: float,
    A_over_At: float,
    T_wall: float,
) -> float:
    """
    Local convective heat transfer coefficient from the Bartz (1957) equation.

    All quantities in SI units. The correlation is:

      h_g = (0.026 / D_t^0.2) * (mu^0.2 * Cp / Pr^0.6)
            * (Pc / c*)^0.8 * (D_t / r_c)^0.1 * (At/A)^0.9 * sigma

    Parameters
    ----------
    D_throat : float
        Throat diameter [m]
    r_c : float
        Throat wall radius of curvature [m]
    Pc_Pa : float
        Chamber pressure [Pa]
    c_star : float
        Characteristic velocity [m/s]
    mu_chamber : float
        Dynamic viscosity at chamber (stagnation) conditions [Pa·s]
    Cp_chamber : float
        Specific heat at chamber conditions [J/(kg·K)]
    Pr_chamber : float
        Prandtl number at chamber conditions [-]
    T_chamber : float
        Chamber (stagnation) temperature [K]
    gamma : float
        Specific heat ratio [-]
    M : float
        Local Mach number [-]
    A_over_At : float
        Local area ratio A/At [-]
    T_wall : float
        Wall temperature estimate [K] (used in sigma correction)

    Returns
    -------
    float
        Convective heat transfer coefficient h_g [W/(m²·K)]
    """
    sigma = bartz_sigma(T_wall, T_chamber, M, gamma)

    h_g = (
        (0.026 / D_throat ** 0.2)
        * (mu_chamber ** 0.2 * Cp_chamber / Pr_chamber ** 0.6)
        * (Pc_Pa / c_star) ** 0.8
        * (D_throat / r_c) ** 0.1
        * (1.0 / A_over_At) ** 0.9
        * sigma
    )
    return h_g


def recovery_temperature(
    T_chamber: float, M: float, gamma: float, Pr: float
) -> float:
    """
    Adiabatic wall temperature (recovery temperature) of the hot gas.

    For turbulent boundary layers, the recovery factor r = Pr^(1/3).

    T_aw = T_chamber * (1 + r*(gamma-1)/2*M^2) / (1 + (gamma-1)/2*M^2)

    Parameters
    ----------
    T_chamber : float
        Chamber (stagnation) temperature [K]
    M : float
        Local Mach number [-]
    gamma : float
        Specific heat ratio [-]
    Pr : float
        Prandtl number [-]

    Returns
    -------
    float
        Adiabatic wall temperature T_aw [K]
    """
    r = Pr ** (1.0 / 3.0)                       # Turbulent recovery factor
    mach_term = (gamma - 1.0) / 2.0 * M ** 2
    T_aw = T_chamber * (1.0 + r * mach_term) / (1.0 + mach_term)
    return T_aw


def compute_hot_gas_profile(
    nozzle: NozzleGeometry,
    gamma: float,
    Pc_Pa: float,
    c_star: float,
    mu_ch: float,
    Cp_ch: float,
    Pr_ch: float,
    T_chamber: float,
    T_wall_guess: float = 1000.0,
) -> dict:
    """
    Compute the hot-gas-side heat transfer coefficient and adiabatic wall
    temperature at every axial station along the nozzle.

    Parameters
    ----------
    nozzle : NozzleGeometry
        Nozzle geometry object
    gamma : float
        Specific heat ratio (use chamber value as representative)
    Pc_Pa : float
        Chamber pressure [Pa]
    c_star : float
        Characteristic velocity [m/s]
    mu_ch : float
        Viscosity at chamber conditions [Pa·s]
    Cp_ch : float
        Specific heat at chamber conditions [J/(kg·K)]
    Pr_ch : float
        Prandtl number at chamber conditions [-]
    T_chamber : float
        Chamber temperature [K]
    T_wall_guess : float
        Initial wall temperature estimate for sigma [K]

    Returns
    -------
    dict with numpy arrays:
        x    : axial positions [m]
        M    : Mach number [-]
        T_aw : adiabatic wall temperature [K]
        h_g  : convective heat transfer coefficient [W/(m²·K)]
        AR   : area ratio A/At [-]
        D    : local diameter [m]
    """
    x, M = compute_mach_profile(nozzle, gamma)
    AR = nozzle.area_ratio(x)
    D  = nozzle.diameter(x)

    T_aw = np.array([
        recovery_temperature(T_chamber, Mi, gamma, Pr_ch) for Mi in M
    ])

    h_g = np.array([
        bartz_coefficient(
            D_throat    = nozzle.D_throat,
            r_c         = nozzle.r_c,
            Pc_Pa       = Pc_Pa,
            c_star      = c_star,
            mu_chamber  = mu_ch,
            Cp_chamber  = Cp_ch,
            Pr_chamber  = Pr_ch,
            T_chamber   = T_chamber,
            gamma       = gamma,
            M           = Mi,
            A_over_At   = ARi,
            T_wall      = T_wall_guess,
        )
        for Mi, ARi in zip(M, AR)
    ])

    return {
        "x":    x,
        "M":    M,
        "T_aw": T_aw,
        "h_g":  h_g,
        "AR":   AR,
        "D":    D,
    }
