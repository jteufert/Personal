#!/usr/bin/env python3
"""
Quick sanity-check for the NRLMSISE-00 calculator.
Compare a few known reference values from Picone et al. (2002) Table 1.

Run:  python test_nrlmsise.py
"""

from nrlmsise_calc import calculate_density


def check(label, alt_km, f107, expected_kg_m3, tol_pct=20):
    rho = calculate_density(alt_km, f107, f107a=f107, ap=4,
                            doy=172, year=2001, lat=45, lon=0, lst=12)
    err = abs(rho - expected_kg_m3) / expected_kg_m3 * 100
    status = "OK" if err < tol_pct else "FAIL"
    print(f"[{status}] {label:40s}  got {rho:.3e} kg/m^3  (ref {expected_kg_m3:.3e})  err {err:.1f}%")


if __name__ == "__main__":
    print("NRLMSISE-00 sanity checks (reference: US Standard Atm / model literature)")
    print("=" * 78)

    # Reference values are approximate; 20% tolerance allows for date/location variance
    check("Sea level (0 km), F10.7=150",          0,    150, 1.225e+0)
    check("Tropopause (~11 km), F10.7=150",       11,   150, 3.64e-1)
    check("Stratopause (~50 km), F10.7=150",      50,   150, 1.03e-3)
    check("Mesopause (~86 km), F10.7=150",        86,   150, 6.0e-6)
    check("Thermosphere 200 km, F10.7=70",       200,    70, 2.5e-10)
    check("Thermosphere 200 km, F10.7=150",      200,   150, 5.0e-10)
    check("Thermosphere 400 km, F10.7=150",      400,   150, 3.0e-12)
    check("Thermosphere 800 km, F10.7=150",      800,   150, 1.0e-14)

    print()
    print("F10.7 sensitivity at 400 km:")
    for f107 in [70, 100, 150, 200, 250]:
        rho = calculate_density(400, f107, f107a=f107, ap=4)
        print(f"  F10.7 = {f107:3d}  ->  {rho:.3e} kg/m^3")
