"""
CEA Interface for Film Cooling Analysis

Wraps rocketcea to provide combustion and transport properties in SI units
at chamber, throat, and exit for GOX/Ethanol propellants.

Unit conversions from CEA English units to SI:
  Temperature:          °R  -> K    (multiply by 5/9)
  Pressure:             psia -> Pa   (multiply by 6894.76)
  Velocity (c*):        ft/s -> m/s  (multiply by 0.3048)
  Viscosity:            millipoise -> Pa·s  (multiply by 1e-4)
  Thermal conductivity: mcal/(cm·s·K) -> W/(m·K)  (multiply by 0.4184)
  Specific heat:        cal/(g·K) -> J/(kg·K)  (multiply by 4186.8)
"""

import numpy as np
from rocketcea.cea_obj import CEA_Obj


# Unit conversion constants
_R_TO_K = 5.0 / 9.0           # Rankine to Kelvin
_PSIA_TO_PA = 6894.76          # psia to Pascal
_FPS_TO_MPS = 0.3048           # ft/s to m/s
_MP_TO_PAS = 1e-4              # millipoise to Pa·s
_MCAL_CMSK_TO_WMK = 0.4184    # mcal/(cm·s·K) to W/(m·K)
_CALGK_TO_JKGK = 4186.8       # cal/(g·K) to J/(kg·K)
_BAR_TO_PSIA = 14.5038         # bar to psia


class CEAInterface:
    """
    Interface to NASA CEA via rocketcea for GOX/Ethanol combustion analysis.

    All inputs and outputs are in SI units unless otherwise noted.
    """

    def __init__(self, ox_name: str = "GOX", fuel_name: str = "Ethanol"):
        self.ox_name = ox_name
        self.fuel_name = fuel_name
        self._cea = CEA_Obj(oxName=ox_name, fuelName=fuel_name)

    def _psia(self, Pc_bar: float) -> float:
        return Pc_bar * _BAR_TO_PSIA

    def _transport_to_si(self, transport_tuple: tuple) -> dict:
        """Convert a (Cp, visc, cond, Pr) transport tuple to SI dict."""
        Cp_cal, visc_mp, cond_mcal, Pr = transport_tuple
        return {
            "cp":  Cp_cal  * _CALGK_TO_JKGK,
            "mu":  visc_mp * _MP_TO_PAS,
            "k":   cond_mcal * _MCAL_CMSK_TO_WMK,
            "Pr":  float(Pr),
        }

    def get_combustion_properties(
        self, Pc_bar: float, MR: float, eps: float
    ) -> dict:
        """
        Return key combustion properties at chamber conditions.

        Parameters
        ----------
        Pc_bar : float
            Chamber pressure [bar]
        MR : float
            Oxidizer-to-fuel mass flow ratio (O/F)
        eps : float
            Nozzle exit area ratio Ae/At

        Returns
        -------
        dict with keys:
            T_c    : Chamber temperature [K]
            gamma  : Specific heat ratio at chamber [-]
            c_star : Characteristic velocity [m/s]
            Isp    : Specific impulse (vacuum) [s]
            MW     : Mean molecular weight of combustion products [g/mol]
            cp, mu, k, Pr : Transport properties (SI)
        """
        Pc = self._psia(Pc_bar)

        T_c_R = self._cea.get_Tcomb(Pc=Pc, MR=MR)
        MW, gamma = self._cea.get_Chamber_MolWt_gamma(Pc=Pc, MR=MR, eps=eps)
        c_star_fps = self._cea.get_Cstar(Pc=Pc, MR=MR)
        Isp = self._cea.get_Isp(Pc=Pc, MR=MR, eps=eps)

        transport = self._transport_to_si(
            self._cea.get_Chamber_Transport(Pc=Pc, MR=MR, eps=eps)
        )

        return {
            "T_c":    float(T_c_R) * _R_TO_K,
            "gamma":  float(gamma),
            "c_star": float(c_star_fps) * _FPS_TO_MPS,
            "Isp":    float(Isp),
            "MW":     float(MW),
            **transport,
        }

    def get_throat_properties(self, Pc_bar: float, MR: float, eps: float) -> dict:
        """
        Return flow properties at the nozzle throat.

        Returns
        -------
        dict with keys: T, gamma, cp, mu, k, Pr (all SI)
        """
        Pc = self._psia(Pc_bar)

        temps_R = self._cea.get_Temperatures(Pc=Pc, MR=MR, eps=eps)
        T_throat_K = float(temps_R[1]) * _R_TO_K

        MW_t, gamma_t = self._cea.get_Throat_MolWt_gamma(Pc=Pc, MR=MR)
        transport = self._transport_to_si(
            self._cea.get_Throat_Transport(Pc=Pc, MR=MR, eps=eps)
        )

        return {
            "T":     T_throat_K,
            "gamma": float(gamma_t),
            **transport,
        }

    def get_exit_properties(self, Pc_bar: float, MR: float, eps: float) -> dict:
        """
        Return flow properties at the nozzle exit.

        Returns
        -------
        dict with keys: T, gamma, cp, mu, k, Pr (all SI)
        """
        Pc = self._psia(Pc_bar)

        temps_R = self._cea.get_Temperatures(Pc=Pc, MR=MR, eps=eps)
        T_exit_K = float(temps_R[2]) * _R_TO_K

        MW_e, gamma_e = self._cea.get_exit_MolWt_gamma(Pc=Pc, MR=MR, eps=eps)
        transport = self._transport_to_si(
            self._cea.get_Exit_Transport(Pc=Pc, MR=MR, eps=eps)
        )

        return {
            "T":     T_exit_K,
            "gamma": float(gamma_e),
            **transport,
        }

    def get_all_properties(self, Pc_bar: float, MR: float, eps: float) -> dict:
        """
        Return combustion properties at chamber, throat, and exit in one call.

        Returns
        -------
        dict with keys 'chamber', 'throat', 'exit', each a property dict.
        """
        return {
            "chamber": self.get_combustion_properties(Pc_bar, MR, eps),
            "throat":  self.get_throat_properties(Pc_bar, MR, eps),
            "exit":    self.get_exit_properties(Pc_bar, MR, eps),
        }

    def print_summary(self, Pc_bar: float, MR: float, eps: float) -> None:
        """Print a formatted summary of combustion and nozzle properties."""
        props = self.get_all_properties(Pc_bar, MR, eps)
        ch = props["chamber"]
        th = props["throat"]
        ex = props["exit"]

        print(f"\n{'='*60}")
        print(f"  CEA Results: {self.ox_name} / {self.fuel_name}")
        print(f"  Pc = {Pc_bar:.2f} bar | O/F = {MR:.3f} | eps = {eps:.1f}")
        print(f"{'='*60}")
        print(f"  {'Property':<30} {'Chamber':>10} {'Throat':>10} {'Exit':>10}")
        print(f"  {'-'*60}")
        print(f"  {'Temperature [K]':<30} {ch['T_c']:>10.0f} {th['T']:>10.0f} {ex['T']:>10.0f}")
        print(f"  {'Gamma [-]':<30} {ch['gamma']:>10.4f} {th['gamma']:>10.4f} {ex['gamma']:>10.4f}")
        print(f"  {'Cp [J/(kg·K)]':<30} {ch['cp']:>10.0f} {th['cp']:>10.0f} {ex['cp']:>10.0f}")
        print(f"  {'Viscosity [uPa·s]':<30} {ch['mu']*1e6:>10.2f} {th['mu']*1e6:>10.2f} {ex['mu']*1e6:>10.2f}")
        print(f"  {'Therm. cond. [mW/(m·K)]':<30} {ch['k']*1e3:>10.1f} {th['k']*1e3:>10.1f} {ex['k']*1e3:>10.1f}")
        print(f"  {'Prandtl [-]':<30} {ch['Pr']:>10.4f} {th['Pr']:>10.4f} {ex['Pr']:>10.4f}")
        print(f"{'='*60}")
        print(f"  Characteristic velocity c* : {ch['c_star']:.1f} m/s")
        print(f"  Specific impulse (vac) Isp : {ch['Isp']:.1f} s")
        print(f"  Mean molecular weight       : {ch['MW']:.2f} g/mol")
        print(f"{'='*60}\n")
