"""
Nozzle Geometry and Isentropic Flow Relations

Generates the axial profile of a conical rocket nozzle (converging-diverging)
and computes local Mach number, temperature, and pressure at each station
using isentropic gas dynamics.

Nozzle sections (axial coordinate x, origin at throat):
  Chamber   : x in [-L_c - L_conv, -L_conv]
  Convergent: x in [-L_conv, 0]  (throat at x=0)
  Divergent : x in [0, L_div]
"""

import numpy as np
from scipy.optimize import brentq
from dataclasses import dataclass, field
from typing import Tuple


@dataclass
class NozzleGeometry:
    """
    Conical nozzle geometry parameterized by throat diameter and area ratio.

    Parameters
    ----------
    D_throat : float
        Throat diameter [m]
    eps : float
        Exit area ratio Ae/At
    Dc_Dt : float
        Chamber-to-throat diameter ratio (typically 2.5-4)
    L_chamber : float
        Chamber cylindrical length [m]
    alpha_conv_deg : float
        Convergent half-angle [deg] (typically 20-45)
    alpha_div_deg : float
        Divergent half-angle [deg] (typically 10-20 for conical)
    rc_over_Rt : float
        Throat curvature radius / throat radius (for Bartz equation, typically 0.5-1.5)
    n_points : int
        Number of axial stations for the profile
    """
    D_throat: float
    eps: float
    Dc_Dt: float = 3.0
    L_chamber: float = 0.05
    alpha_conv_deg: float = 30.0
    alpha_div_deg: float = 15.0
    rc_over_Rt: float = 0.8
    n_points: int = 300

    # Derived attributes (computed in __post_init__)
    R_throat: float = field(init=False)
    R_chamber: float = field(init=False)
    R_exit: float = field(init=False)
    L_conv: float = field(init=False)
    L_div: float = field(init=False)
    r_c: float = field(init=False)   # Throat radius of curvature
    x: np.ndarray = field(init=False, repr=False)

    def __post_init__(self):
        self.R_throat  = self.D_throat / 2.0
        self.R_chamber = self.R_throat * self.Dc_Dt
        self.R_exit    = self.R_throat * np.sqrt(self.eps)

        alpha_conv = np.radians(self.alpha_conv_deg)
        alpha_div  = np.radians(self.alpha_div_deg)

        self.L_conv = (self.R_chamber - self.R_throat) / np.tan(alpha_conv)
        self.L_div  = (self.R_exit - self.R_throat)   / np.tan(alpha_div)
        self.r_c    = self.rc_over_Rt * self.R_throat

        # Build axial coordinate array (throat at x = 0)
        x_chamber  = np.linspace(-(self.L_chamber + self.L_conv),
                                  -self.L_conv, self.n_points // 5)
        x_conv     = np.linspace(-self.L_conv, 0.0, self.n_points // 3)
        x_div      = np.linspace(0.0, self.L_div, self.n_points // 2)
        self.x = np.unique(np.concatenate([x_chamber, x_conv, x_div]))

    @property
    def total_length(self) -> float:
        """Total nozzle length from chamber head to exit [m]."""
        return self.L_chamber + self.L_conv + self.L_div

    def radius(self, x: np.ndarray) -> np.ndarray:
        """
        Local wall radius as a function of axial position [m].

        x < -L_conv           : chamber (cylindrical)
        -L_conv <= x <= 0     : convergent cone
        x > 0                 : divergent cone
        """
        x = np.asarray(x, dtype=float)
        r = np.empty_like(x)

        alpha_conv = np.radians(self.alpha_conv_deg)
        alpha_div  = np.radians(self.alpha_div_deg)

        # Chamber section
        mask_ch = x <= -self.L_conv
        r[mask_ch] = self.R_chamber

        # Convergent section
        mask_cv = (~mask_ch) & (x <= 0.0)
        r[mask_cv] = self.R_throat + (-x[mask_cv]) * np.tan(alpha_conv)

        # Divergent section
        mask_dv = x > 0.0
        r[mask_dv] = self.R_throat + x[mask_dv] * np.tan(alpha_div)

        return r

    def area(self, x: np.ndarray) -> np.ndarray:
        """Local cross-sectional area [m²]."""
        return np.pi * self.radius(x) ** 2

    def area_ratio(self, x: np.ndarray) -> np.ndarray:
        """Local area ratio A/At."""
        A_t = np.pi * self.R_throat ** 2
        return self.area(x) / A_t

    def diameter(self, x: np.ndarray) -> np.ndarray:
        """Local diameter [m]."""
        return 2.0 * self.radius(x)

    def print_summary(self) -> None:
        """Print a table of key nozzle dimensions."""
        A_t = np.pi * self.R_throat ** 2
        A_e = np.pi * self.R_exit ** 2
        print(f"\n{'='*50}")
        print(f"  Nozzle Geometry")
        print(f"{'='*50}")
        print(f"  Throat diameter  : {self.D_throat*1000:.2f} mm")
        print(f"  Chamber diameter : {2*self.R_chamber*1000:.2f} mm")
        print(f"  Exit diameter    : {2*self.R_exit*1000:.2f} mm")
        print(f"  Area ratio (eps) : {self.eps:.1f}")
        print(f"  Throat area      : {A_t*1e6:.4f} mm²")
        print(f"  Chamber length   : {self.L_chamber*1000:.1f} mm")
        print(f"  Convergent length: {self.L_conv*1000:.1f} mm")
        print(f"  Divergent length : {self.L_div*1000:.1f} mm")
        print(f"  Throat curv. rad.: {self.r_c*1000:.2f} mm")
        print(f"  Conv. half-angle : {self.alpha_conv_deg:.1f} deg")
        print(f"  Div.  half-angle : {self.alpha_div_deg:.1f} deg")
        print(f"{'='*50}\n")


# ---------------------------------------------------------------------------
# Isentropic flow utilities
# ---------------------------------------------------------------------------

def area_ratio_from_mach(M: float, gamma: float) -> float:
    """
    Isentropic area ratio A/A* as a function of Mach number.

    A/A* = (1/M) * [(2/(gamma+1)) * (1 + (gamma-1)/2 * M^2)]^((gamma+1)/(2*(gamma-1)))
    """
    t = 1.0 + (gamma - 1.0) / 2.0 * M ** 2
    exp = (gamma + 1.0) / (2.0 * (gamma - 1.0))
    return (1.0 / M) * (2.0 / (gamma + 1.0) * t) ** exp


def mach_from_area_ratio(
    A_over_At: float, gamma: float, supersonic: bool = True
) -> float:
    """
    Solve for Mach number given the isentropic area ratio A/A*.

    Parameters
    ----------
    A_over_At : float
        Area ratio A/A* (must be >= 1)
    gamma : float
        Specific heat ratio
    supersonic : bool
        If True, return the supersonic solution (M > 1);
        if False, return the subsonic solution (0 < M < 1).

    Returns
    -------
    float
        Mach number
    """
    if A_over_At < 1.0:
        raise ValueError(f"Area ratio must be >= 1 (got {A_over_At:.4f})")
    if A_over_At == 1.0:
        return 1.0

    f = lambda M: area_ratio_from_mach(M, gamma) - A_over_At

    if supersonic:
        # Supersonic root: M > 1
        # Upper bracket: find M_high such that f(M_high) > 0
        M_high = 2.0
        while f(M_high) < 0:
            M_high *= 2.0
        return brentq(f, 1.0 + 1e-9, M_high, xtol=1e-8)
    else:
        # Subsonic root: 0 < M < 1
        return brentq(f, 1e-9, 1.0 - 1e-9, xtol=1e-8)


def isentropic_temperature(T0: float, M: float, gamma: float) -> float:
    """
    Local static temperature from stagnation temperature.

    T = T0 / (1 + (gamma-1)/2 * M^2)
    """
    return T0 / (1.0 + (gamma - 1.0) / 2.0 * M ** 2)


def isentropic_pressure(P0: float, M: float, gamma: float) -> float:
    """
    Local static pressure from stagnation (chamber) pressure.

    P = P0 / (1 + (gamma-1)/2 * M^2)^(gamma/(gamma-1))
    """
    return P0 / (1.0 + (gamma - 1.0) / 2.0 * M ** 2) ** (
        gamma / (gamma - 1.0)
    )


def compute_mach_profile(
    nozzle: NozzleGeometry, gamma: float
) -> Tuple[np.ndarray, np.ndarray]:
    """
    Compute Mach number at every axial station in the nozzle.

    Returns
    -------
    x : np.ndarray
        Axial positions [m] (throat at x=0)
    M : np.ndarray
        Local Mach number [-]
    """
    x = nozzle.x
    AR = nozzle.area_ratio(x)
    M = np.empty_like(x)

    for i, (xi, ar) in enumerate(zip(x, AR)):
        if abs(xi) < 1e-10:
            M[i] = 1.0
        elif xi < 0:
            # Subsonic (convergent section or chamber)
            if ar <= 1.0 + 1e-6:
                M[i] = 1.0
            else:
                M[i] = mach_from_area_ratio(ar, gamma, supersonic=False)
        else:
            # Supersonic (divergent section)
            M[i] = mach_from_area_ratio(ar, gamma, supersonic=True)

    return x, M
