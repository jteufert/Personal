"""
Film Cooling Model

Models the protection of the nozzle wall by a film of coolant injected
at the start of the convergent section. Two complementary approaches:

1. Energy Balance (ODE integration) - physically tracks the film temperature
   as heat from the hot gas raises it along the wall. The film is
   considered "spent" when it reaches the local adiabatic wall temperature.

2. Goldstein Effectiveness Correlation - empirical curve-fit for turbulent
   slot-injection film cooling effectiveness decay:

     eta(x) = 1 / (1 + beta * xi^n)

   where xi = (x/s) / M_film, s = effective slot height, M_film = blowing ratio.
   Constants: beta = 0.329, n = 0.8 for turbulent 2D slot injection (Goldstein 1971).

The energy balance approach is used as the primary model because it directly
accounts for the local heat transfer coefficient h_g(x) and nozzle geometry.

Reference:
  Goldstein, R.J. (1971). Film Cooling. Advances in Heat Transfer, Vol. 7.
  Huzel & Huang (1992). Modern Engineering for Design of LPREs. AIAA.
"""

import numpy as np
from scipy.integrate import solve_ivp


def film_slot_height(
    m_dot_film: float,
    rho_film: float,
    v_film: float,
    D_injection: float,
) -> float:
    """
    Effective slot height for a circumferential annular film injection.

    s = m_dot_film / (rho_film * v_film * pi * D_inj)

    Parameters
    ----------
    m_dot_film : float
        Film coolant mass flow [kg/s]
    rho_film : float
        Film coolant density at injection [kg/m³]
    v_film : float
        Film injection velocity [m/s]
    D_injection : float
        Diameter at the injection point [m]

    Returns
    -------
    float
        Effective slot height s [m]
    """
    return m_dot_film / (rho_film * v_film * np.pi * D_injection)


def blowing_ratio(
    rho_film: float,
    v_film: float,
    rho_gas: float,
    v_gas: float,
) -> float:
    """
    Blowing ratio (injection momentum flux ratio).

    M_B = (rho_film * v_film) / (rho_gas * v_gas)
    """
    return (rho_film * v_film) / (rho_gas * v_gas)


def goldstein_effectiveness(
    x_from_injection: np.ndarray,
    slot_height: float,
    M_blowing: float,
    beta: float = 0.329,
    n: float = 0.8,
) -> np.ndarray:
    """
    Goldstein (1971) turbulent slot-injection film cooling effectiveness.

    eta = 1 / (1 + beta * xi^n)
    xi  = x / (s * M_blowing)

    Parameters
    ----------
    x_from_injection : np.ndarray
        Distance from injection point [m]
    slot_height : float
        Effective slot height s [m]
    M_blowing : float
        Blowing ratio M_B [-]
    beta : float
        Correlation coefficient (default 0.329 for turbulent slot)
    n : float
        Decay exponent (default 0.8 for turbulent slot)

    Returns
    -------
    np.ndarray
        Film cooling effectiveness eta [-], 1 at injection, decays toward 0.
    """
    x = np.asarray(x_from_injection, dtype=float)
    x = np.clip(x, 0.0, None)   # effectiveness only defined downstream

    denom = slot_height * max(M_blowing, 1e-6)
    xi = x / denom
    eta = 1.0 / (1.0 + beta * xi ** n)
    return eta


def energy_balance_film(
    x_stations: np.ndarray,
    h_g: np.ndarray,
    T_aw: np.ndarray,
    D_wall: np.ndarray,
    m_dot_film: float,
    Cp_film: float,
    T_film_in: float,
    x_injection: float,
) -> dict:
    """
    Integrate the film energy balance ODE to track film temperature.

    The film absorbs heat from the hot gas via convection at coefficient h_g,
    heating from its inlet temperature until it reaches the local recovery
    temperature (effectiveness -> 0).

    ODE:
      m_dot_film * Cp_film * dT_film/dx = h_g(x) * (T_aw(x) - T_film(x)) * pi * D(x)

    Parameters
    ----------
    x_stations : np.ndarray
        Axial positions [m] (full nozzle, throat at x=0)
    h_g : np.ndarray
        Convective heat transfer coefficient [W/(m²·K)] at each station
    T_aw : np.ndarray
        Adiabatic wall temperature [K] at each station (no film)
    D_wall : np.ndarray
        Local wall diameter [m] at each station
    m_dot_film : float
        Film coolant mass flow [kg/s]
    Cp_film : float
        Film coolant specific heat [J/(kg·K)]
    T_film_in : float
        Film coolant inlet temperature [K]
    x_injection : float
        Axial position of film injection [m]

    Returns
    -------
    dict with numpy arrays:
        T_film    : Film temperature at each station [K]
                    (equals T_film_in upstream of injection)
        T_aw_fc   : Effective adiabatic wall temp with film cooling [K]
        eta       : Film cooling effectiveness [-]
        active    : Boolean mask where the film is thermally active
    """
    # Identify stations downstream of injection
    active_mask = x_stations >= x_injection

    # Interpolation functions for h_g and T_aw between CEA stations
    from scipy.interpolate import interp1d
    h_g_fn  = interp1d(x_stations, h_g,  kind="linear", fill_value="extrapolate")
    T_aw_fn = interp1d(x_stations, T_aw, kind="linear", fill_value="extrapolate")
    D_fn    = interp1d(x_stations, D_wall, kind="linear", fill_value="extrapolate")

    # ODE right-hand side
    def dT_film_dx(x, T_f):
        hg   = max(h_g_fn(x), 0.0)
        Taw  = T_aw_fn(x)
        D    = max(D_fn(x), 1e-6)
        dT   = hg * (Taw - T_f[0]) * np.pi * D / (m_dot_film * Cp_film)
        return [dT]

    # Integration domain: from injection point to exit
    x_active = x_stations[active_mask]
    if len(x_active) < 2:
        # No active film region
        T_film = np.full_like(x_stations, T_film_in)
        eta    = np.zeros_like(x_stations)
        T_aw_fc = T_aw.copy()
        return {
            "T_film": T_film,
            "T_aw_fc": T_aw_fc,
            "eta": eta,
            "active": active_mask,
        }

    sol = solve_ivp(
        dT_film_dx,
        t_span=(x_active[0], x_active[-1]),
        y0=[T_film_in],
        t_eval=x_active,
        method="RK45",
        rtol=1e-6,
        atol=1e-8,
    )

    T_film_active = sol.y[0]

    # Build full-length T_film array
    T_film = np.full_like(x_stations, T_film_in, dtype=float)
    T_film[active_mask] = T_film_active

    # Effectiveness: eta = (T_aw - T_film) / (T_aw - T_film_in)
    # Clipped to [0, 1] to handle numerical overshoot
    denom = T_aw - T_film_in
    with np.errstate(invalid="ignore", divide="ignore"):
        eta_raw = np.where(
            np.abs(denom) > 1.0,
            (T_aw - T_film) / denom,
            0.0,
        )
    eta = np.clip(eta_raw, 0.0, 1.0)

    # Upstream of injection: no film cooling
    eta[~active_mask] = 0.0

    # Effective adiabatic wall temperature with film cooling
    # T_aw_fc = T_film + (1 - eta) * (T_aw - T_film_in)
    # Equivalently: T_aw_fc = eta * T_film_in + (1-eta) * T_aw
    T_aw_fc = eta * T_film_in + (1.0 - eta) * T_aw

    return {
        "T_film":   T_film,
        "T_aw_fc":  T_aw_fc,
        "eta":      eta,
        "active":   active_mask,
    }


def mass_flow_from_thrust(
    thrust_N: float,
    Isp_s: float,
    g0: float = 9.80665,
) -> float:
    """
    Total propellant mass flow from thrust and specific impulse.

    m_dot = Thrust / (Isp * g0)
    """
    return thrust_N / (Isp_s * g0)


def split_mass_flows(
    m_dot_total: float,
    MR: float,
    film_fraction_of_fuel: float,
) -> dict:
    """
    Split total mass flow into oxidizer, core fuel, and film coolant flows.

    Parameters
    ----------
    m_dot_total : float
        Total propellant mass flow [kg/s]
    MR : float
        O/F mixture ratio
    film_fraction_of_fuel : float
        Film coolant flow as a fraction of total fuel flow (0 to 1)

    Returns
    -------
    dict with:
        m_dot_ox   : Oxidizer mass flow [kg/s]
        m_dot_fuel : Total fuel mass flow [kg/s]
        m_dot_film : Film coolant mass flow [kg/s]
        m_dot_core : Core (combustion) mass flow [kg/s]
    """
    m_dot_ox   = m_dot_total * MR / (1.0 + MR)
    m_dot_fuel = m_dot_total / (1.0 + MR)
    m_dot_film = m_dot_fuel * film_fraction_of_fuel
    m_dot_core = m_dot_total - m_dot_film

    return {
        "m_dot_ox":   m_dot_ox,
        "m_dot_fuel": m_dot_fuel,
        "m_dot_film": m_dot_film,
        "m_dot_core": m_dot_core,
    }
