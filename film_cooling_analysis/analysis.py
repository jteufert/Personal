"""
Integrated Film Cooling Thermal Analysis

Combines CEA combustion data, Bartz heat transfer, film cooling ODE,
and radiation cooling into a single steady-state wall temperature solver.

At each axial station the thermal equilibrium is:

  q_in  = h_g * (T_aw_fc - T_wall)        [hot gas with film cooling]
  q_out = epsilon * sigma * (T_wall^4 - T_bg^4)  [radiation to environment]

Setting q_in = q_out yields a nonlinear equation for T_wall(x) at each
station, solved with scipy's brentq root finder.

For stations where the film is inactive (upstream of injection or after
film burnout), T_aw_fc -> T_aw (no film benefit) and wall temperatures
are typically limited only by radiation.
"""

import numpy as np
from scipy.optimize import brentq
from typing import Optional

from .cea_interface import CEAInterface
from .nozzle import NozzleGeometry
from .heat_transfer import compute_hot_gas_profile
from .film_cooling import (
    energy_balance_film,
    mass_flow_from_thrust,
    split_mass_flows,
)
from .radiation import radiation_flux, SIGMA_SB


class ThrusterAnalysis:
    """
    Full thermal analysis of a film-cooled, radiatively cooled thruster.

    Parameters
    ----------
    config : dict
        Configuration dictionary with the following keys:

        Propellant:
          ox_name          : str   - Oxidizer name for CEA (default 'GOX')
          fuel_name        : str   - Fuel name for CEA (default 'Ethanol')

        Operating conditions:
          Pc               : float - Chamber pressure [bar]
          MR               : float - O/F mixture ratio
          thrust           : float - Target thrust [N] (vacuum)

        Nozzle geometry:
          eps              : float - Exit area ratio Ae/At
          Dc_Dt            : float - Chamber/throat diameter ratio
          L_chamber        : float - Chamber cylindrical length [m]
          alpha_conv_deg   : float - Convergent half-angle [deg]
          alpha_div_deg    : float - Divergent half-angle [deg]
          rc_over_Rt       : float - Throat curvature / throat radius
          n_points         : int   - Number of axial analysis stations

        Film cooling:
          film_fraction    : float - Film flow / total fuel flow (0-1)
          Cp_film          : float - Film coolant Cp [J/(kg·K)]
          T_film_in        : float - Film inlet temperature [K]

        Wall properties:
          emissivity       : float - Wall emissivity [-]
          T_wall_max       : float - Material temperature limit [K]

        Environment:
          T_background     : float - Radiation sink temperature [K]
    """

    def __init__(self, config: dict):
        self.cfg = {
            # Defaults for a small GOX/Ethanol thruster
            "ox_name":        "GOX",
            "fuel_name":      "Ethanol",
            "Pc":             10.0,
            "MR":             1.4,
            # Sizing: set D_throat [m] to fix geometry; or set thrust [N]
            # to size the throat from performance. D_throat takes priority.
            "D_throat":       None,
            "thrust":         20.0,
            "eps":            50.0,
            "Dc_Dt":          3.0,
            "L_chamber":      0.05,
            "alpha_conv_deg": 30.0,
            "alpha_div_deg":  15.0,
            "rc_over_Rt":     0.8,
            "n_points":       400,
            "film_fraction":  0.10,
            "Cp_film":        2570.0,
            "T_film_in":      300.0,
            "emissivity":     0.85,
            "T_wall_max":     1644.0,    # 1371 °C / 1500 °F
            "T_background":   4.0,
        }
        self.cfg.update(config)

        # Results storage (populated by run())
        self.cea_props: Optional[dict] = None
        self.nozzle: Optional[NozzleGeometry] = None
        self.results: Optional[dict] = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def run(self) -> dict:
        """
        Execute the full analysis pipeline.

        Steps:
          1. CEA: combustion and transport properties
          2. Nozzle sizing from thrust
          3. Bartz heat transfer profile
          4. Mass flow splits (oxidizer / fuel / film)
          5. Film cooling ODE integration
          6. Wall temperature solution at each station
          7. Assemble and return results dict

        Returns
        -------
        dict
            All computed profiles keyed by variable name.
        """
        cfg = self.cfg

        # --- Step 1: CEA ------------------------------------------------
        cea = CEAInterface(cfg["ox_name"], cfg["fuel_name"])
        all_props = cea.get_all_properties(cfg["Pc"], cfg["MR"], cfg["eps"])
        ch = all_props["chamber"]
        self.cea_props = all_props

        # --- Step 2: Nozzle sizing ---------------------------------------
        g0    = 9.80665
        Pc_Pa = cfg["Pc"] * 1e5   # bar -> Pa

        if cfg.get("D_throat") is not None:
            # Fixed hardware: throat diameter given, thrust is an output
            D_throat = float(cfg["D_throat"])
            A_throat = np.pi * (D_throat / 2.0) ** 2
            m_dot_total = Pc_Pa * A_throat / ch["c_star"]
            thrust_N = m_dot_total * ch["Isp"] * g0
        else:
            # Design to thrust target: size the throat
            m_dot_total = mass_flow_from_thrust(cfg["thrust"], ch["Isp"], g0)
            A_throat    = m_dot_total * ch["c_star"] / Pc_Pa
            D_throat    = 2.0 * np.sqrt(A_throat / np.pi)
            thrust_N    = cfg["thrust"]

        nozzle = NozzleGeometry(
            D_throat      = D_throat,
            eps           = cfg["eps"],
            Dc_Dt         = cfg["Dc_Dt"],
            L_chamber     = cfg["L_chamber"],
            alpha_conv_deg= cfg["alpha_conv_deg"],
            alpha_div_deg = cfg["alpha_div_deg"],
            rc_over_Rt    = cfg["rc_over_Rt"],
            n_points      = cfg["n_points"],
        )
        self.nozzle = nozzle

        # --- Step 3: Hot gas heat transfer profile -----------------------
        hg_prof = compute_hot_gas_profile(
            nozzle     = nozzle,
            gamma      = ch["gamma"],
            Pc_Pa      = Pc_Pa,
            c_star     = ch["c_star"],
            mu_ch      = ch["mu"],
            Cp_ch      = ch["cp"],
            Pr_ch      = ch["Pr"],
            T_chamber  = ch["T_c"],
        )
        x    = hg_prof["x"]
        M    = hg_prof["M"]
        T_aw = hg_prof["T_aw"]
        h_g  = hg_prof["h_g"]
        AR   = hg_prof["AR"]
        D    = hg_prof["D"]

        # --- Step 4: Mass flow splits ------------------------------------
        flows = split_mass_flows(m_dot_total, cfg["MR"], cfg["film_fraction"])
        m_dot_film = flows["m_dot_film"]

        # Film injection at the start of the convergent section
        x_injection = -nozzle.L_conv

        # --- Step 5: Film cooling ----------------------------------------
        film = energy_balance_film(
            x_stations   = x,
            h_g          = h_g,
            T_aw         = T_aw,
            D_wall       = D,
            m_dot_film   = m_dot_film,
            Cp_film      = cfg["Cp_film"],
            T_film_in    = cfg["T_film_in"],
            x_injection  = x_injection,
        )
        T_film   = film["T_film"]
        T_aw_fc  = film["T_aw_fc"]
        eta      = film["eta"]

        # --- Step 6: Wall temperature ------------------------------------
        eps_wall = cfg["emissivity"]
        T_bg     = cfg["T_background"]

        T_wall_fc   = self._solve_wall_temp_array(h_g, T_aw_fc, eps_wall, T_bg)
        T_wall_nofc = self._solve_wall_temp_array(h_g, T_aw,    eps_wall, T_bg)

        # --- Step 7: Heat flux profiles ----------------------------------
        q_conv_fc   = h_g * (T_aw_fc - T_wall_fc)
        q_conv_nofc = h_g * (T_aw - T_wall_nofc)

        # --- Assemble results --------------------------------------------
        self.results = {
            # Axial profile
            "x":           x,
            "M":           M,
            "AR":          AR,
            "D":           D,
            # Gas temperatures
            "T_aw":        T_aw,
            "T_aw_fc":     T_aw_fc,
            "T_film":      T_film,
            # Film effectiveness
            "eta":         eta,
            # Heat transfer
            "h_g":         h_g,
            "q_conv_fc":   q_conv_fc,
            "q_conv_nofc": q_conv_nofc,
            # Wall temperatures
            "T_wall_fc":   T_wall_fc,
            "T_wall_nofc": T_wall_nofc,
            # Thruster sizing
            "D_throat":    D_throat,
            "m_dot_total": m_dot_total,
            "m_dot_film":  m_dot_film,
            "m_dot_core":  flows["m_dot_core"],
            "m_dot_ox":    flows["m_dot_ox"],
            "m_dot_fuel":  flows["m_dot_fuel"],
            # Performance
            "thrust_N":    thrust_N,
            "Isp":         ch["Isp"],
            "c_star":      ch["c_star"],
            "T_c":         ch["T_c"],
        }
        return self.results

    def print_summary(self) -> None:
        """Print a human-readable summary after run() has been called."""
        if self.results is None:
            raise RuntimeError("Call run() before print_summary().")

        r = self.results
        cfg = self.cfg
        nozzle = self.nozzle

        print(f"\n{'='*65}")
        print(f"  Film Cooling Analysis: {cfg['ox_name']} / {cfg['fuel_name']}")
        print(f"{'='*65}")
        Pc_psi = cfg["Pc"] * 14.5038
        thrust_label = ("computed" if cfg.get("D_throat") else "target")
        print(f"  Operating Conditions")
        print(f"    Chamber pressure  : {cfg['Pc']:.3f} bar  ({Pc_psi:.1f} psi)")
        print(f"    Mixture ratio O/F : {cfg['MR']:.3f}")
        print(f"    Thrust ({thrust_label:8s}) : {r['thrust_N']:.2f} N (vacuum)")
        print(f"    Specific impulse  : {r['Isp']:.1f} s")
        print(f"    Chamber temp.     : {r['T_c']:.0f} K")
        print(f"    Char. velocity c* : {r['c_star']:.1f} m/s")
        print()
        print(f"  Nozzle Geometry")
        print(f"    Throat diameter   : {r['D_throat']*1e3:.3f} mm")
        print(f"    Chamber diameter  : {nozzle.D_throat*nozzle.Dc_Dt*1e3:.2f} mm")
        print(f"    Exit diameter     : {2*nozzle.R_exit*1e3:.2f} mm")
        print(f"    Area ratio eps    : {cfg['eps']:.1f}")
        print(f"    Total length      : {nozzle.total_length*1e3:.1f} mm")
        print()
        print(f"  Mass Flows")
        print(f"    Total propellant  : {r['m_dot_total']*1e3:.3f} g/s")
        print(f"    Oxidizer          : {r['m_dot_ox']*1e3:.3f} g/s")
        print(f"    Fuel (total)      : {r['m_dot_fuel']*1e3:.3f} g/s")
        print(f"    Film coolant      : {r['m_dot_film']*1e3:.3f} g/s "
              f"({cfg['film_fraction']*100:.1f}% of fuel)")
        print(f"    Core combustion   : {r['m_dot_core']*1e3:.3f} g/s")
        print()
        print(f"  Thermal Results")
        # Throat index (closest to x=0)
        i_th = int(np.argmin(np.abs(r['x'])))
        print(f"    Peak h_g (throat) : {r['h_g'][i_th]/1e3:.1f} kW/(m²·K)")
        print(f"    Peak heat flux    : {r['q_conv_nofc'].max()/1e6:.2f} MW/m² (no film)")
        print(f"    Peak heat flux    : {r['q_conv_fc'].max()/1e6:.2f} MW/m² (with film)")
        print(f"    Peak T_aw         : {r['T_aw'].max():.0f} K (no film, at chamber)")
        print(f"    T_aw at throat    : {r['T_aw'][i_th]:.0f} K")
        print(f"    T_aw_fc at throat : {r['T_aw_fc'][i_th]:.0f} K (with film)")
        print(f"    Peak wall temp    : {r['T_wall_fc'].max():.0f} K  (film+rad cooling)")
        print(f"    Peak wall temp    : {r['T_wall_nofc'].max():.0f} K  (radiation only)")
        print(f"    Material limit    : {cfg['T_wall_max']:.0f} K")

        peak_fc = r['T_wall_fc'].max()
        if peak_fc > cfg["T_wall_max"]:
            excess = peak_fc - cfg["T_wall_max"]
            print(f"    *** WARNING: Wall exceeds limit by {excess:.0f} K ***")
        else:
            margin = cfg["T_wall_max"] - peak_fc
            print(f"    Temperature margin: {margin:.0f} K  [OK]")
        print(f"{'='*65}\n")

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _solve_wall_temp(
        h_g: float,
        T_aw_fc: float,
        emissivity: float,
        T_bg: float,
    ) -> float:
        """
        Solve for steady-state wall temperature at a single axial station.

        Equation: h_g*(T_aw_fc - T_w) = eps*sigma*(T_w^4 - T_bg^4)
        Rearranged to f(T_w) = 0 and solved with brentq.
        """
        def residual(T_w):
            q_in  = h_g * (T_aw_fc - T_w)
            q_out = emissivity * SIGMA_SB * (T_w ** 4 - T_bg ** 4)
            return q_in - q_out

        # The wall temperature must be between T_bg and T_aw_fc
        T_lo = max(T_bg + 1.0, 1.0)
        T_hi = max(T_aw_fc - 1.0, T_lo + 1.0)

        # Make sure we bracket a root
        if residual(T_lo) * residual(T_hi) > 0:
            # If no sign change, the radiation can't balance convection:
            # wall heats to near T_aw_fc (very thin film or low emissivity)
            return T_hi

        return brentq(residual, T_lo, T_hi, xtol=0.1)

    def _solve_wall_temp_array(
        self,
        h_g: np.ndarray,
        T_aw_fc: np.ndarray,
        emissivity: float,
        T_bg: float,
    ) -> np.ndarray:
        """Solve wall temperature at every axial station."""
        return np.array([
            self._solve_wall_temp(hgi, tawi, emissivity, T_bg)
            for hgi, tawi in zip(h_g, T_aw_fc)
        ])
