#!/usr/bin/env python3
"""
NRLMSISE-00 Atmospheric Density Calculator
Called from Excel VBA via Shell() to write results to a temp file.

Usage (command line):
    python nrlmsise_calc.py <alt_km> <f107> [f107a] [ap] [doy] [lat] [lon] [lst]

Output (stdout):
    density_kg_m3

Requires: msise00  (pip install msise00)
"""

import sys
import os
import json
from datetime import datetime, date, timedelta


def calculate_density(alt_km, f107, f107a=None, ap=4.0,
                      doy=172, year=2024, lat=45.0, lon=0.0, lst=12.0):
    """
    Calculate atmospheric total mass density using NRLMSISE-00.

    Parameters
    ----------
    alt_km : float
        Geodetic altitude in km
    f107 : float
        Daily F10.7 solar flux index (solar flux units, 1 SFU = 10^-22 W/m^2/Hz)
    f107a : float, optional
        81-day average F10.7. Defaults to f107 if not provided.
    ap : float, optional
        Geomagnetic Ap index. Default 4 (quiet conditions).
    doy : int, optional
        Day of year (1-366). Default 172 (near summer solstice).
    year : int, optional
        Year. Default 2024.
    lat : float, optional
        Geodetic latitude in degrees. Default 45 N.
    lon : float, optional
        Longitude in degrees. Default 0 (Greenwich).
    lst : float, optional
        Local solar time in hours. Default 12 (noon).

    Returns
    -------
    float
        Total mass density in kg/m^3
    """
    if f107a is None:
        f107a = f107

    # Build datetime from year + doy
    base = date(year, 1, 1)
    d = base + timedelta(days=int(doy) - 1)
    hour = int(lst)
    minute = int((lst - hour) * 60)
    t = datetime(d.year, d.month, d.day, hour, minute)

    try:
        import msise00
        # msise00.run returns an xarray.Dataset
        atmos = msise00.run(t, float(alt_km), float(lat), float(lon),
                            float(f107a), float(f107), float(ap))
        # 'Total' is in kg/m^3
        density = float(atmos["Total"].values.ravel()[0])
        return density
    except ImportError:
        pass

    try:
        from nrlmsise00 import msise_flat
        # nrlmsise00 (scholte) interface
        # returns array: [d0-d8, T_low, T_exo] where d5=total density (g/cm^3)
        sec = lst * 3600.0
        out = msise_flat(doy, sec, float(alt_km), float(lat), float(lon),
                         float(lst), float(f107a), float(f107), float(ap))
        # out[5] = total mass density in g/cm^3 → convert to kg/m^3
        density = out[5] * 1000.0
        return density
    except ImportError:
        pass

    raise RuntimeError(
        "No NRLMSISE-00 package found.\n"
        "Install one of:\n"
        "  pip install msise00\n"
        "  pip install nrlmsise00"
    )


if __name__ == "__main__":
    args = sys.argv[1:]
    if len(args) < 2:
        print("ERROR: Usage: python nrlmsise_calc.py <alt_km> <f107> "
              "[f107a] [ap] [doy] [lat] [lon] [lst]", file=sys.stderr)
        sys.exit(1)

    try:
        alt_km = float(args[0])
        f107   = float(args[1])
        f107a  = float(args[2]) if len(args) > 2 else None
        ap     = float(args[3]) if len(args) > 3 else 4.0
        doy    = int(args[4])   if len(args) > 4 else 172
        lat    = float(args[5]) if len(args) > 5 else 45.0
        lon    = float(args[6]) if len(args) > 6 else 0.0
        lst    = float(args[7]) if len(args) > 7 else 12.0

        density = calculate_density(alt_km, f107, f107a, ap, doy, 2024, lat, lon, lst)
        # Print result so VBA can read it from a temp file
        print(f"{density:.6e}")

    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
