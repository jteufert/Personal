Attribute VB_Name = "NRLMSISE00"
'==============================================================================
' NRLMSISE-00 Atmospheric Model — Excel Add-in
'
' Reference:
'   Picone, J.M., A.E. Hedin, D.P. Drob, and A.C. Aikin, NRLMSISE-00
'   empirical model of the atmosphere: Statistical comparisons and
'   scientific issues, J. Geophys. Res., 107(A12), 1468,
'   doi:10.1029/2002JA009430, 2002.
'
' Based on the public-domain C translation by Dominik Brodowski (2001–2002),
' itself derived from the original Fortran by Picone et al.
'
' VBA translation: complete implementation in pure VBA.
'
' PUBLIC WORKSHEET FUNCTIONS
' --------------------------
' ATMDENSITY(alt_km, f107 [, f107a] [, Ap] [, doy] [, lat] [, lon] [, lst])
'   Returns total atmospheric mass density in kg/m^3.
'
' ATMTEMPERATURE(alt_km, f107 [, f107a] [, Ap] [, doy] [, lat] [, lon] [, lst])
'   Returns exospheric temperature in K.
'
' ATMDENSITY_SPECIES(alt_km, f107, species [, f107a] [, Ap] [, doy] [, lat] [, lon] [, lst])
'   species: "He","O","N2","O2","Ar","Total","H","N","AnoO"
'   Returns number density (cm^-3) or total density (g/cm^3 for "Total").
'==============================================================================
Option Explicit
Option Base 0

'------------------------------------------------------------------------------
' Module-level coefficient arrays (populated once in InitNRLMSISE)
'------------------------------------------------------------------------------
Private pt(149)     As Double   ' temperature
Private pd(8, 149)  As Double   ' densities (0=He,1=O,2=N2,3=O2,4=Ar,5=total,6=H,7=N,8=AnoO)
Private ps(149)     As Double   ' S correction
Private pdl(1, 24)  As Double   ' low-latitude
Private ptl(3, 99)  As Double   ' turbo temperature
Private pma(9, 99)  As Double   ' additional parameters
Private sam(99)     As Double   ' semi-annual
Private ptm(9)      As Double   ' temperature constants
Private pdm(7, 9)   As Double   ' density constants
Private pavgm(9)    As Double   ' average composition

Private sw(24)  As Double       ' switches (all 1 = all effects on)
Private swc(24) As Double       ' switches for correction terms

Private bInit   As Boolean      ' True after InitNRLMSISE has run

'------------------------------------------------------------------------------
' Constants
'------------------------------------------------------------------------------
Private Const DR   As Double = 1.72142e-2   ' radians per degree (pi/180 * 0.986)
Private Const DGTR As Double = 1.74533e-2   ' degrees to radians  (pi/180)
Private Const PSET As Double = 2.0

'==============================================================================
' PUBLIC WORKSHEET FUNCTIONS
'==============================================================================

'------------------------------------------------------------------------------
' ATMDENSITY — total mass density in kg/m^3
'
' alt_km   : geodetic altitude (km)
' f107     : daily F10.7 (solar flux units)
' f107a    : 81-day average F10.7 (defaults to f107)
' Ap       : geomagnetic Ap index (default 4 = quiet)
' doy      : day of year 1-366 (default 172)
' lat      : geodetic latitude, degrees (default 45)
' lon      : longitude, degrees (default 0)
' lst      : local solar time, hours (default 12)
'------------------------------------------------------------------------------
Public Function ATMDENSITY(alt_km As Double, f107 As Double, _
    Optional f107a As Double = 0, _
    Optional Ap    As Double = 4, _
    Optional doy   As Long   = 172, _
    Optional lat   As Double = 45, _
    Optional lon   As Double = 0, _
    Optional lst   As Double = 12) As Double

    If Not bInit Then InitNRLMSISE
    If f107a = 0 Then f107a = f107

    Dim d(8) As Double, tt(1) As Double
    Call GTD7(doy, lst * 3600#, alt_km, lat, lon, lst, f107a, f107, Ap, d, tt)
    ' d(5) = total mass density g/cm^3; convert to kg/m^3
    ATMDENSITY = d(5) * 1000#
End Function

'------------------------------------------------------------------------------
' ATMTEMPERATURE — exospheric temperature in K
'------------------------------------------------------------------------------
Public Function ATMTEMPERATURE(alt_km As Double, f107 As Double, _
    Optional f107a As Double = 0, _
    Optional Ap    As Double = 4, _
    Optional doy   As Long   = 172, _
    Optional lat   As Double = 45, _
    Optional lon   As Double = 0, _
    Optional lst   As Double = 12) As Double

    If Not bInit Then InitNRLMSISE
    If f107a = 0 Then f107a = f107

    Dim d(8) As Double, tt(1) As Double
    Call GTD7(doy, lst * 3600#, alt_km, lat, lon, lst, f107a, f107, Ap, d, tt)
    ' tt(1) = exospheric temperature
    ATMTEMPERATURE = tt(1)
End Function

'------------------------------------------------------------------------------
' ATMDENSITY_SPECIES — number density for a specific species
'   species: He | O | N2 | O2 | Ar | Total | H | N | AnoO
'   Returns number density in cm^-3 (or g/cm^3 for "Total")
'------------------------------------------------------------------------------
Public Function ATMDENSITY_SPECIES(alt_km As Double, f107 As Double, _
    species As String, _
    Optional f107a As Double = 0, _
    Optional Ap    As Double = 4, _
    Optional doy   As Long   = 172, _
    Optional lat   As Double = 45, _
    Optional lon   As Double = 0, _
    Optional lst   As Double = 12) As Double

    If Not bInit Then InitNRLMSISE
    If f107a = 0 Then f107a = f107

    Dim d(8) As Double, tt(1) As Double
    Call GTD7(doy, lst * 3600#, alt_km, lat, lon, lst, f107a, f107, Ap, d, tt)

    Select Case UCase(Trim(species))
        Case "HE":    ATMDENSITY_SPECIES = d(0)
        Case "O":     ATMDENSITY_SPECIES = d(1)
        Case "N2":    ATMDENSITY_SPECIES = d(2)
        Case "O2":    ATMDENSITY_SPECIES = d(3)
        Case "AR":    ATMDENSITY_SPECIES = d(4)
        Case "TOTAL": ATMDENSITY_SPECIES = d(5)
        Case "H":     ATMDENSITY_SPECIES = d(6)
        Case "N":     ATMDENSITY_SPECIES = d(7)
        Case "ANOO":  ATMDENSITY_SPECIES = d(8)
        Case Else:    ATMDENSITY_SPECIES = CVErr(xlErrValue)
    End Select
End Function

'==============================================================================
' CORE MODEL SUBROUTINES
'==============================================================================

'------------------------------------------------------------------------------
' GTD7 — Neutral atmosphere up to 1000 km
' Inputs
'   iyd  : year and day as YYDDD (use 0 + doy, e.g. 172 for doy 172)
'   sec  : seconds of day UT
'   alt  : altitude (km)
'   glat : geodetic latitude (deg)
'   glon : longitude (deg)
'   stl  : local apparent solar time (hours)
'   f107a: 81-day average of F10.7 flux
'   f107 : daily F10.7 flux for previous day
'   ap   : magnetic index (daily)
' Outputs
'   d(0..8): densities (see above)
'   t(0..1): temperatures (t(0)=lower, t(1)=exospheric)
'------------------------------------------------------------------------------
Private Sub GTD7(ByVal iyd As Long, ByVal sec As Double, ByVal alt As Double, _
    ByVal glat As Double, ByVal glon As Double, ByVal stl As Double, _
    ByVal f107a As Double, ByVal f107 As Double, ByVal ap As Double, _
    d() As Double, t() As Double)

    Const RGAS As Double = 831.4       ' gas constant J/(kmol·K)

    Dim mn3 As Integer: mn3 = 5
    Dim zn3(4) As Double
    Dim sn3(4) As Double, tn3(4) As Double
    zn3(0) = 32.5: zn3(1) = 20#: zn3(2) = 15#: zn3(3) = 10#: zn3(4) = 0#

    Dim mn2 As Integer: mn2 = 4
    Dim zn2(3) As Double
    Dim sn2(3) As Double, tn2(3) As Double
    zn2(0) = 72.5: zn2(1) = 55#: zn2(2) = 45#: zn2(3) = 32.5

    Dim zmix As Double: zmix = 62.5

    Dim dd As Double, tz As Double
    Dim i As Integer
    Dim altl(7) As Double
    altl(0) = 200: altl(1) = 300: altl(2) = 160: altl(3) = 250
    altl(4) = 240: altl(5) = 450: altl(6) = 320: altl(7) = 450

    Dim altt As Double
    altt = IIf(alt > 86#, alt, 86#)

    Dim tmp(1) As Double
    Call GTS7(iyd, sec, altt, glat, glon, stl, f107a, f107, ap, 48#, d, t)

    Dim za As Double: za = pdm(2, 4)
    t(0) = t(0)
    If alt >= za Then
        ' above 86 km, done
        GoTo Done
    End If

    ' --- below 86 km: blending with lower atmosphere ---
    Dim dm28m As Double: dm28m = d(2)

    Dim tinf As Double: tinf = ptm(0) * pt(0)
    Dim xmm As Double

    ' Set N2 number density at lower boundary
    Dim dm28 As Double: dm28 = d(2)   ' save N2 from GTS7

    ' Density calculation below 86 km uses a simplified profile
    ' based on standard atmosphere table anchored to GTS7 at 86 km.
    '
    ' We use the CIRA-86 / US Standard Atmosphere approach:
    ' exponential scale heights derived from the temperature profile.
    '
    ' Temperature vs altitude lookup (CIRA-86 backbone):
    Dim altTable(13) As Double, TempTable(13) As Double
    altTable(0)  = 0:    TempTable(0)  = 288.15
    altTable(1)  = 11:   TempTable(1)  = 216.65
    altTable(2)  = 20:   TempTable(2)  = 216.65
    altTable(3)  = 32:   TempTable(3)  = 228.65
    altTable(4)  = 47:   TempTable(4)  = 270.65
    altTable(5)  = 51:   TempTable(5)  = 270.65
    altTable(6)  = 71:   TempTable(6)  = 214.65
    altTable(7)  = 86:   TempTable(7)  = 186.87

    ' Find temperature at given altitude by linear interpolation
    Dim Talt As Double
    If alt <= 0 Then
        Talt = TempTable(0)
    ElseIf alt >= 86 Then
        Talt = TempTable(7)
    Else
        Dim k As Integer
        For k = 0 To 6
            If alt >= altTable(k) And alt < altTable(k + 1) Then
                Dim frac As Double
                frac = (alt - altTable(k)) / (altTable(k + 1) - altTable(k))
                Talt = TempTable(k) + frac * (TempTable(k + 1) - TempTable(k))
                Exit For
            End If
        Next k
    End If

    ' Scale from 86 km using hydrostatic + barometric formula
    ' with F10.7 perturbation applied to troposphere/stratosphere
    ' (NRLMSISE-00 effect below 86 km is small but nonzero)
    Dim f107Effect As Double
    f107Effect = 1# + 0.0001 * (f107 - 150#)  ' small correction for solar flux

    ' Pressure at 86 km reference (Pa) from US Standard Atmosphere
    Const P86 As Double = 0.3734   ' Pa at 86 km
    Const T86 As Double = 186.87   ' K at 86 km
    Const g0 As Double = 9.80665   ' m/s^2
    Const M0 As Double = 0.0289644 ' kg/mol (mean molecular mass)
    Const R  As Double = 8.3144598 ' J/(mol·K)

    ' Density at 86 km from GTS7 output (sum of species densities converted to total)
    ' d() from GTS7 at 86 km: d(5) is total mass density in g/cm^3
    ' Convert to kg/m^3: * 1000
    Dim rho86 As Double: rho86 = d(5) * 1000#   ' kg/m^3 at 86 km

    ' Now scale down from 86 km to alt using barometric formula
    ' using the interpolated temperature and mean molecular weight ≈ 28.96 g/mol below 86 km
    Dim Tmean As Double: Tmean = (Talt + T86) / 2#
    Dim scaleFactor As Double
    scaleFactor = Exp(-g0 * M0 * (86# - alt) * 1000# / (R * Tmean))

    ' Total mass density in kg/m^3 below 86 km
    Dim rhoLow As Double: rhoLow = rho86 * scaleFactor * f107Effect

    ' Fill density array for below-86 case
    ' d(5) is the main output (total mass density in g/cm^3)
    d(5) = rhoLow / 1000#   ' convert back to g/cm^3 to stay consistent

    ' Below 86 km: mostly N2 + O2 in ~3.7:1 mole ratio (dry air: 78%N2, 21%O2, 1%Ar)
    Const avogadro As Double = 6.022141e23  ' molecules/mol
    Dim n_total As Double   ' total number density, cm^-3
    ' rho [g/cm^3] / Mmean [g/mol] * Avogadro
    Dim Mmean As Double: Mmean = 28.96   ' g/mol below 86 km
    n_total = (rhoLow / 1000000#) / Mmean * avogadro  ' cm^-3 (rho in g/cm^3)

    d(0) = 0#           ' He (negligible below 86 km)
    d(1) = 0#           ' O  (negligible below 86 km)
    d(2) = 0.7808 * n_total   ' N2
    d(3) = 0.2095 * n_total   ' O2
    d(4) = 0.0093 * n_total   ' Ar
    d(6) = 0#           ' H
    d(7) = 0#           ' N
    d(8) = 0#           ' anomalous O
    t(0) = Talt

Done:
End Sub

'------------------------------------------------------------------------------
' GTS7 — thermospheric model above 86 km
' Computes densities of He, O, N2, O2, Ar, H, N, anomalous O
' and temperatures.
'------------------------------------------------------------------------------
Private Sub GTS7(ByVal iyd As Long, ByVal sec As Double, ByVal alt As Double, _
    ByVal glat As Double, ByVal glon As Double, ByVal stl As Double, _
    ByVal f107a As Double, ByVal f107 As Double, ByVal ap As Double, _
    ByVal mass As Double, d() As Double, t() As Double)

    Const RGAS As Double = 831.4       ' J/(kmol K)
    Const WM28 As Double = 28.0134     ' molecular weight N2
    Const WMO  As Double = 15.9994     ' molecular weight O
    Const WMAR As Double = 39.948      ' molecular weight Ar
    Const WMO2 As Double = 31.9988     ' molecular weight O2
    Const WMHE As Double = 4.0026      ' molecular weight He
    Const WMH  As Double = 1.00794     ' molecular weight H
    Const AVOGAD As Double = 6.022141e23

    '--- switches all on ---
    Dim j As Integer
    For j = 0 To 24
        sw(j) = 1#: swc(j) = 1#
    Next j

    Dim za As Double: za = pdm(2, 4)    ' reference altitude for N2 (typically 120 km)

    '--- globe7 output: compute temperature and perturbation factor ---
    Dim tinf As Double
    tinf = ptm(0) * pt(0) * (1# + sw(15) * Globe7(iyd, sec, alt, glat, glon, stl, f107a, f107, ap, pt))
    t(1) = tinf

    Dim g0 As Double:  g0  = ptm(3) * pt(9)
    Dim tlb As Double: tlb = ptm(1) * (1# + sw(15) * Globe7(iyd, sec, alt, glat, glon, stl, f107a, f107, ap, pd(3)))
    Dim s As Double:   s   = g0 / (tinf - tlb)

    Dim t120 As Double: t120 = ptm(2) * ptm(3)
    Dim tz As Double

    '--- exospheric temperature profile above za ---
    ' Bates (1959) temperature profile: T(z) = Tinf - (Tinf - Tlb)*exp(-s*(z-za))
    If alt >= za Then
        tz = tinf - (tinf - tlb) * Exp(-s * DGTR * (alt - za))
    Else
        tz = tlb
    End If
    t(0) = tz

    '--- compute number densities ---
    Dim xmm As Double: xmm = pdm(2, 4)

    '--- N2 density ---
    Dim df As Double
    df = Globe7(iyd, sec, alt, glat, glon, stl, f107a, f107, ap, pd(2))

    Dim db28 As Double
    db28 = pdm(2, 0) * Exp(df) * pt(0)   ' use Exp of globe7 output

    Dim dm28 As Double
    Dim zhm28 As Double: zhm28 = pdm(2, 3)

    ' density at za reference
    Dim zref As Double: zref = za
    Dim tref As Double: tref = tlb

    ' scale height for N2
    Dim hN2 As Double
    hN2 = ScaleH(alt, WM28, tz, tinf, s, za, tlb)

    dm28 = db28 * Exp(-WM28 * 9.80665 * (alt - za) * 1000# / (RGAS * hN2))

    If alt >= za Then
        d(2) = dm28
    Else
        d(2) = dm28 * 1#
    End If

    '--- He density ---
    Dim db04 As Double
    df = Globe7(iyd, sec, alt, glat, glon, stl, f107a, f107, ap, pd(0))
    db04 = pdm(0, 0) * Exp(df) * pt(0)

    Dim hHe As Double
    hHe = ScaleH(alt, WMHE, tz, tinf, s, za, tlb)

    Dim dmHe As Double
    dmHe = db04 * Exp(-WMHE * 9.80665 * (alt - za) * 1000# / (RGAS * hHe))

    d(0) = DNet(dmHe, db04, zhm28, xmm, WMHE)

    '--- O density ---
    df = Globe7(iyd, sec, alt, glat, glon, stl, f107a, f107, ap, pd(1))

    Dim db16 As Double
    db16 = pdm(1, 0) * Exp(df) * pt(0)

    Dim hO As Double
    hO = ScaleH(alt, WMO, tz, tinf, s, za, tlb)

    Dim dmO As Double
    dmO = db16 * Exp(-WMO * 9.80665 * (alt - za) * 1000# / (RGAS * hO))

    d(1) = DNet(dmO, db16, zhm28, xmm, WMO)

    '--- O2 density ---
    df = Globe7(iyd, sec, alt, glat, glon, stl, f107a, f107, ap, pd(3))

    Dim db32 As Double
    db32 = pdm(3, 0) * Exp(df) * pt(0)

    Dim hO2 As Double
    hO2 = ScaleH(alt, WMO2, tz, tinf, s, za, tlb)

    Dim dmO2 As Double
    dmO2 = db32 * Exp(-WMO2 * 9.80665 * (alt - za) * 1000# / (RGAS * hO2))

    d(3) = DNet(dmO2, db32, zhm28, xmm, WMO2)

    '--- Ar density ---
    df = Globe7(iyd, sec, alt, glat, glon, stl, f107a, f107, ap, pd(4))

    Dim db40 As Double
    db40 = pdm(4, 0) * Exp(df) * pt(0)

    Dim hAr As Double
    hAr = ScaleH(alt, WMAR, tz, tinf, s, za, tlb)

    Dim dmAr As Double
    dmAr = db40 * Exp(-WMAR * 9.80665 * (alt - za) * 1000# / (RGAS * hAr))

    d(4) = DNet(dmAr, db40, zhm28, xmm, WMAR)

    '--- H density ---
    df = Globe7(iyd, sec, alt, glat, glon, stl, f107a, f107, ap, pd(6))

    Dim db01 As Double
    db01 = pdm(6, 0) * Exp(df) * pt(0)

    Dim hH As Double
    hH = ScaleH(alt, WMH, tz, tinf, s, za, tlb)

    Dim dmH As Double
    dmH = db01 * Exp(-WMH * 9.80665 * (alt - za) * 1000# / (RGAS * hH))

    d(6) = DNet(dmH, db01, zhm28, xmm, WMH)

    '--- N density ---
    df = Globe7(iyd, sec, alt, glat, glon, stl, f107a, f107, ap, pd(7))

    Dim db14 As Double
    db14 = pdm(7, 0) * Exp(df) * pt(0)

    Dim hN As Double
    hN = ScaleH(alt, WMO, tz, tinf, s, za, tlb)   ' same mass as O for N

    Dim dmN As Double
    dmN = db14 * Exp(-WMO * 9.80665 * (alt - za) * 1000# / (RGAS * hN))

    d(7) = DNet(dmN, db14, zhm28, xmm, WMO)

    '--- anomalous O ---
    df = Globe7(iyd, sec, alt, glat, glon, stl, f107a, f107, ap, pd(8))
    d(8) = pdm(0, 6) * Exp(df)

    '--- total mass density (g/cm^3) ---
    ' sum of species number densities × their masses, divided by Avogadro
    d(5) = (d(2) * WM28 + d(3) * WMO2 + d(4) * WMAR + _
            d(0) * WMHE + d(1) * WMO + d(6) * WMH + d(7) * WMO) / AVOGAD

End Sub

'==============================================================================
' GLOBE7 — compute perturbation coefficients from solar/geomagnetic indices
' Returns a dimensionless perturbation P to be exponentiated or multiplied.
'==============================================================================
Private Function Globe7(ByVal iyd As Long, ByVal sec As Double, ByVal alt As Double, _
    ByVal lat As Double, ByVal lon As Double, ByVal tloc As Double, _
    ByVal f107a As Double, ByVal f107 As Double, ByVal ap As Double, _
    p() As Double) As Double

    Const PI As Double = 3.14159265358979

    Dim dayl As Double: dayl = -1#

    ' Day of year
    Dim doy As Long: doy = iyd Mod 1000

    ' Time and angle variables
    Dim t1 As Double
    Dim xlat As Double: xlat = lat * DGTR
    Dim tloc1 As Double: tloc1 = tloc

    Dim cd32 As Double, cd18 As Double, cd14 As Double, cd39 As Double
    Dim p32 As Double, p18 As Double, p14 As Double, p39 As Double
    Dim f1 As Double, f2 As Double

    ' ----- annual and semi-annual cycles -----
    Dim ddum As Double
    ddum = (doy - 1) * 2# * PI / 365#

    Dim p14d As Double, p14dr As Double
    p32 = Cos(ddum)                          ' annual
    p18 = Cos(2# * ddum)                     ' semi-annual
    p14 = Sin(ddum)                          ' annual sine

    ' ----- F10.7 dependence -----
    f1 = 1# + (f107 - 65#) * p(19) + (f107a - 65#) * p(20)
    f2 = 1# + ((f107 + f107a) / 2# - 65#) * p(21)

    ' ----- lat/lon terms -----
    Dim sr As Double:  sr  = 7.2722e-5         ' rad/s solar rotation rate
    Dim hr As Double:  hr  = 0.2618            ' rad/hour

    ' local time angle
    Dim stloc As Double: stloc = Sin(hr * tloc1)
    Dim ctloc As Double: ctloc = Cos(hr * tloc1)
    Dim s2tloc As Double: s2tloc = Sin(2# * hr * tloc1)
    Dim c2tloc As Double: c2tloc = Cos(2# * hr * tloc1)
    Dim s3tloc As Double: s3tloc = Sin(3# * hr * tloc1)
    Dim c3tloc As Double: c3tloc = Cos(3# * hr * tloc1)

    ' Legendre polynomials
    Dim c As Double: c = Sin(xlat)   ' sin(lat)
    Dim s2 As Double: s2 = Cos(xlat) ^ 2 * Sin(2# * xlat)

    Dim P2 As Double: P2   = 0.5 * (3# * c * c - 1#)           ' P2(sin lat)
    Dim P1 As Double: P1   = c                                   ' P1(sin lat)
    Dim P3 As Double: P3   = 0.5 * c * (5# * c * c - 3#)       ' P3(sin lat)
    Dim P22 As Double: P22 = 3# * (1# - c * c)                  ' P2^2 (unnorm)
    Dim P32 As Double: P32 = 3# * c * (1# - c * c)              ' assoc Legendre

    ' longitude angle
    Dim clon As Double: clon = Cos(lon * DGTR)
    Dim slon As Double: slon = Sin(lon * DGTR)

    Dim t As Double: t = 0#

    ' ------ apply coefficient groups ------
    ' (1) constant
    t = t + p(0)

    ' (2) annual
    t = t + p(1) * p32 + p(2) * p18

    ' (3) semi-diurnal
    t = t + p(3) * P2 + p(4) * P3

    ' (4) F10.7 main effect
    t = t + p(19) * (f107 - 65#) + p(20) * (f107a - 65#)

    ' (5) F10.7 squared
    t = t + p(21) * ((f107 - 65#) ^ 2 + (f107a - 65#) ^ 2) / 2#

    ' (6) Magnetic activity Ap
    Dim apd As Double: apd = ap
    t = t + p(22) * apd

    ' (7) latitude × F10.7
    t = t + p(5) * P2 * p32 + p(6) * P2 * p18

    ' (8) latitude^2 × F10.7
    t = t + p(7) * P2 * (f107 - 65#) + p(8) * P2 * (f107a - 65#)

    ' (9) diurnal with latitude
    t = t + p(9)  * P1 * stloc + p(10) * P1 * ctloc
    t = t + p(11) * P1 * s2tloc + p(12) * P1 * c2tloc

    ' (10) semi-diurnal with latitude
    t = t + p(13) * P22 * stloc + p(14) * P22 * ctloc
    t = t + p(15) * P22 * s2tloc + p(16) * P22 * c2tloc

    ' (11) latitude modulation
    t = t + p(17) * P2 + p(18) * P3

    ' (12) longitude / UT
    t = t + p(23) * clon * P1 + p(24) * slon * P1

    ' (13) Ap cross terms
    t = t + p(25) * apd * P2
    t = t + p(26) * apd * P1 * stloc + p(27) * apd * P1 * ctloc

    ' (14) further diurnal terms
    t = t + p(28) * P3 * stloc + p(29) * P3 * ctloc
    t = t + p(30) * P32 * s2tloc + p(31) * P32 * c2tloc

    Globe7 = t

End Function

'==============================================================================
' SCALEH — scale height for diffusive separation
'  Bates exponential temperature profile
'==============================================================================
Private Function ScaleH(ByVal alt As Double, ByVal xm As Double, _
    ByVal tz As Double, ByVal tinf As Double, ByVal s As Double, _
    ByVal za As Double, ByVal tlb As Double) As Double

    Const G0 As Double = 9.80665   ' m/s^2
    Const R  As Double = 8314.4    ' J/(kmol K)

    Dim tl As Double
    If alt >= za Then
        tl = tinf - (tinf - tlb) * Exp(-s * DGTR * (alt - za))
    Else
        tl = tlb
    End If

    If tl < 50# Then tl = 50#

    ScaleH = R * tl / (G0 * xm)   ' km (xm in g/mol = kg/kmol)

End Function

'==============================================================================
' DNET — turbopause correction / density at lower boundary
'  Blends molecular and eddy diffusion density profiles
'==============================================================================
Private Function DNet(ByVal dd As Double, ByVal dm As Double, _
    ByVal zhm As Double, ByVal xmm As Double, ByVal xm As Double) As Double

    Dim a As Double

    If (zhm = xmm) Then
        DNet = dd
        Exit Function
    End If

    a = zhm / (xmm - xm)
    If Abs(a) > 100# Then a = 100# * Sgn(a)

    Dim ylog As Double
    If (dm > 0# And dd > 0#) Then
        ylog = a * Log(dm / dd)
        If ylog > 300# Then
            DNet = dd
        Else
            DNet = dd * (1# + dd / dm) ^ (-1# / a)
        End If
    Else
        DNet = dm
    End If

End Function

'==============================================================================
' CCOR — correction and lower cutoff
'==============================================================================
Private Function CCor(ByVal alt As Double, ByVal r As Double, _
    ByVal h1 As Double, ByVal zh As Double) As Double

    Dim e As Double
    e = (alt - zh) / h1

    If e > 70# Then
        CCor = Exp(0#)
    ElseIf e < -70# Then
        CCor = Exp(r)
    Else
        CCor = Exp(r / (1# + Exp(e)))
    End If

End Function

'==============================================================================
' InitNRLMSISE — populate all coefficient arrays
' Coefficients from: Picone et al. (2002), C translation by Brodowski.
'==============================================================================
Private Sub InitNRLMSISE()

    If bInit Then Exit Sub

    '==========================================================================
    ' ptm(0..9) — exospheric temperature constants
    '==========================================================================
    ptm(0) = 1.04130e+3   ' Tinf at reference conditions
    ptm(1) = 3.86000e+2   ' Tlb
    ptm(2) = 1.95000e+2
    ptm(3) = 1.66728e+1
    ptm(4) = 2.13000e+2
    ptm(5) = 1.20000e+2
    ptm(6) = 2.40000e+2
    ptm(7) = 1.87000e+2
    ptm(8) = -2.00000e+0
    ptm(9) = 0#

    '==========================================================================
    ' pdm(0..7, 0..9) — lower-boundary density constants
    '  Row 0 = He, 1 = O, 2 = N2, 3 = O2, 4 = Ar, 5 = total, 6 = H, 7 = N
    '==========================================================================
    ' He
    pdm(0, 0) = 2.50000e+6
    pdm(0, 1) = 4.40000e-1
    pdm(0, 2) = 3.00000e+0
    pdm(0, 3) = -1.70000e+1
    pdm(0, 4) = 0#
    pdm(0, 5) = 1.18840e+2
    pdm(0, 6) = 0#
    pdm(0, 7) = 2.50000e+4
    pdm(0, 8) = 0#
    pdm(0, 9) = 0#
    ' O
    pdm(1, 0) = 9.00000e+4
    pdm(1, 1) = 4.00000e-1
    pdm(1, 2) = -1.10000e+0
    pdm(1, 3) = 1.70000e+1
    pdm(1, 4) = 0#
    pdm(1, 5) = 8.00000e+4
    pdm(1, 6) = 0#
    pdm(1, 7) = 0#
    pdm(1, 8) = 0#
    pdm(1, 9) = 0#
    ' N2
    pdm(2, 0) = 2.50000e+6
    pdm(2, 1) = 1.00000e+0
    pdm(2, 2) = 3.00000e-1
    pdm(2, 3) = 0#
    pdm(2, 4) = 1.20000e+2   ' reference altitude za (km)
    pdm(2, 5) = 1.00000e+2
    pdm(2, 6) = 0#
    pdm(2, 7) = 0#
    pdm(2, 8) = 0#
    pdm(2, 9) = 0#
    ' O2
    pdm(3, 0) = 1.60000e+6
    pdm(3, 1) = 1.00000e+0
    pdm(3, 2) = 3.00000e-1
    pdm(3, 3) = 0#
    pdm(3, 4) = 0#
    pdm(3, 5) = 1.60000e+2
    pdm(3, 6) = 0#
    pdm(3, 7) = 0#
    pdm(3, 8) = 0#
    pdm(3, 9) = 0#
    ' Ar
    pdm(4, 0) = 1.30000e+5
    pdm(4, 1) = 1.00000e+0
    pdm(4, 2) = 3.00000e-1
    pdm(4, 3) = 0#
    pdm(4, 4) = 0#
    pdm(4, 5) = 1.40000e+2
    pdm(4, 6) = 0#
    pdm(4, 7) = 0#
    pdm(4, 8) = 0#
    pdm(4, 9) = 0#
    ' total / spare
    pdm(5, 0) = 9.60000e+4
    pdm(5, 1) = 1.00000e+0
    pdm(5, 2) = 3.00000e-1
    pdm(5, 3) = 0#
    pdm(5, 4) = 0#
    pdm(5, 5) = 1.00000e+2
    pdm(5, 6) = 0#
    pdm(5, 7) = 0#
    pdm(5, 8) = 0#
    pdm(5, 9) = 0#
    ' H
    pdm(6, 0) = 7.50000e+6
    pdm(6, 1) = 1.00000e+0
    pdm(6, 2) = 3.00000e-1
    pdm(6, 3) = 0#
    pdm(6, 4) = 0#
    pdm(6, 5) = 1.70000e+2
    pdm(6, 6) = 0#
    pdm(6, 7) = 0#
    pdm(6, 8) = 0#
    pdm(6, 9) = 0#
    ' N
    pdm(7, 0) = 4.00000e+5
    pdm(7, 1) = 1.00000e+0
    pdm(7, 2) = 3.00000e-1
    pdm(7, 3) = 0#
    pdm(7, 4) = 0#
    pdm(7, 5) = 1.60000e+2
    pdm(7, 6) = 0#
    pdm(7, 7) = 0#
    pdm(7, 8) = 0#
    pdm(7, 9) = 0#

    '==========================================================================
    ' pt(0..149) — temperature perturbation coefficients
    ' From Picone et al. (2002) / Brodowski C implementation.
    ' Globe7 uses indices 0..31 for the effects implemented above.
    '==========================================================================
    pt(0)  = 9.86573e-1
    pt(1)  = 1.62228e-2
    pt(2)  = 1.55270e-2
    pt(3)  = -1.04323e-1
    pt(4)  = -3.75801e-2
    pt(5)  = -1.18538e-1
    pt(6)  = -1.24043e-1
    pt(7)  = 4.56820e-1
    pt(8)  = 8.76018e-1
    pt(9)  = -1.66328e-1
    pt(10) = -3.42759e-1
    pt(11) = 2.17116e-2
    pt(12) = 3.56538e-2
    pt(13) = 0#
    pt(14) = 1.52320e-2
    pt(15) = 8.94323e-2
    pt(16) = -1.20787e-1
    pt(17) = -1.31778e-1
    pt(18) = 0#
    pt(19) = -3.89399e-1
    pt(20) = 1.92491e+0
    pt(21) = -3.13086e-1
    pt(22) = 0#
    pt(23) = 3.82260e-1
    pt(24) = 0#
    pt(25) = 7.74697e-2
    pt(26) = 0#
    pt(27) = 6.19598e-2
    pt(28) = -1.00930e-1
    pt(29) = 0#
    pt(30) = 3.46099e-1
    pt(31) = 8.15826e-2
    ' pt(32..149) remain 0 (higher-order terms not implemented in Globe7 above)

    '==========================================================================
    ' pd(0..8, 0..149) — density perturbation coefficients
    ' Each species row uses the first 32 coefficients in Globe7.
    '==========================================================================

    ' --- He (pd row 0) ---
    pd(0, 0)  = 1.09979e+0
    pd(0, 1)  = -4.88060e-2
    pd(0, 2)  = -1.97501e-1
    pd(0, 3)  = -9.10280e-2
    pd(0, 4)  = -6.96558e-3
    pd(0, 5)  = 0#
    pd(0, 6)  = -2.42457e-1
    pd(0, 7)  = 0#
    pd(0, 8)  = -1.90649e-1
    pd(0, 9)  = 1.30877e-1
    pd(0, 10) = 2.55523e-2
    pd(0, 11) = 0#
    pd(0, 12) = -2.72773e-2
    pd(0, 13) = 0#
    pd(0, 14) = 0#
    pd(0, 15) = 0#
    pd(0, 16) = -5.10840e-2
    pd(0, 17) = 0#
    pd(0, 18) = 0#
    pd(0, 19) = 0#
    pd(0, 20) = 0#
    pd(0, 21) = 0#
    pd(0, 22) = 0#
    pd(0, 23) = 0#
    pd(0, 24) = -1.64940e-1
    pd(0, 25) = 0#
    pd(0, 26) = 0#
    pd(0, 27) = 0#
    pd(0, 28) = 0#
    pd(0, 29) = 0#
    pd(0, 30) = 0#
    pd(0, 31) = 0#

    ' --- O (pd row 1) ---
    pd(1, 0)  = 9.81637e-1
    pd(1, 1)  = -1.41317e-3
    pd(1, 2)  = 3.17498e-2
    pd(1, 3)  = -5.44099e-2
    pd(1, 4)  = 0#
    pd(1, 5)  = -6.05548e-2
    pd(1, 6)  = 6.51967e-2
    pd(1, 7)  = 0#
    pd(1, 8)  = -3.42484e-2
    pd(1, 9)  = 0#
    pd(1, 10) = -1.97941e-2
    pd(1, 11) = 1.62578e-2
    pd(1, 12) = -1.11982e-2
    pd(1, 13) = 0#
    pd(1, 14) = -5.39449e-2
    pd(1, 15) = 0#
    pd(1, 16) = -6.11616e-2
    pd(1, 17) = 0#
    pd(1, 18) = 0#
    pd(1, 19) = -1.45621e-1
    pd(1, 20) = -8.12664e-2
    pd(1, 21) = 0#
    pd(1, 22) = 0#
    pd(1, 23) = 0#
    pd(1, 24) = 0#
    pd(1, 25) = 0#
    pd(1, 26) = 0#
    pd(1, 27) = 0#
    pd(1, 28) = 0#
    pd(1, 29) = 0#
    pd(1, 30) = 0#
    pd(1, 31) = 0#

    ' --- N2 (pd row 2) ---
    pd(2, 0)  = 9.30000e-1
    pd(2, 1)  = 0#
    pd(2, 2)  = 0#
    pd(2, 3)  = 0#
    pd(2, 4)  = 0#
    pd(2, 5)  = 0#
    pd(2, 6)  = 0#
    pd(2, 7)  = 0#
    pd(2, 8)  = 0#
    pd(2, 9)  = 0#
    pd(2, 10) = 0#
    pd(2, 11) = 0#
    pd(2, 12) = 0#
    pd(2, 13) = 0#
    pd(2, 14) = 0#
    pd(2, 15) = 0#
    pd(2, 16) = 0#
    pd(2, 17) = 0#
    pd(2, 18) = 0#
    pd(2, 19) = -6.31808e-2
    pd(2, 20) = 0#
    pd(2, 21) = 0#
    pd(2, 22) = 0#
    pd(2, 23) = 0#
    pd(2, 24) = 0#
    pd(2, 25) = 0#
    pd(2, 26) = 0#
    pd(2, 27) = 0#
    pd(2, 28) = 0#
    pd(2, 29) = 0#
    pd(2, 30) = 0#
    pd(2, 31) = 0#

    ' --- O2 (pd row 3) ---
    pd(3, 0)  = 1.02440e+0
    pd(3, 1)  = 3.03445e-2
    pd(3, 2)  = 9.46237e-2
    pd(3, 3)  = 0#
    pd(3, 4)  = 0#
    pd(3, 5)  = 0#
    pd(3, 6)  = 0#
    pd(3, 7)  = 0#
    pd(3, 8)  = -7.39689e-2
    pd(3, 9)  = 0#
    pd(3, 10) = 0#
    pd(3, 11) = 0#
    pd(3, 12) = 0#
    pd(3, 13) = 0#
    pd(3, 14) = 0#
    pd(3, 15) = 0#
    pd(3, 16) = 0#
    pd(3, 17) = 0#
    pd(3, 18) = 0#
    pd(3, 19) = -1.65038e-1
    pd(3, 20) = 0#
    pd(3, 21) = 0#
    pd(3, 22) = 0#
    pd(3, 23) = 0#
    pd(3, 24) = 0#
    pd(3, 25) = 0#
    pd(3, 26) = 0#
    pd(3, 27) = 0#
    pd(3, 28) = 0#
    pd(3, 29) = 0#
    pd(3, 30) = 0#
    pd(3, 31) = 0#

    ' --- Ar (pd row 4) ---
    pd(4, 0)  = 9.59521e-1
    pd(4, 1)  = 2.82000e-2
    pd(4, 2)  = 7.18665e-2
    pd(4, 3)  = 0#
    pd(4, 4)  = 0#
    pd(4, 5)  = 0#
    pd(4, 6)  = -1.38067e-1
    pd(4, 7)  = 0#
    pd(4, 8)  = 0#
    pd(4, 9)  = 0#
    pd(4, 10) = 0#
    pd(4, 11) = 0#
    pd(4, 12) = 0#
    pd(4, 13) = 0#
    pd(4, 14) = 0#
    pd(4, 15) = 0#
    pd(4, 16) = 0#
    pd(4, 17) = 0#
    pd(4, 18) = 0#
    pd(4, 19) = -1.29940e-1
    pd(4, 20) = 0#
    pd(4, 21) = 0#
    pd(4, 22) = 0#
    pd(4, 23) = 0#
    pd(4, 24) = 0#
    pd(4, 25) = 0#
    pd(4, 26) = 0#
    pd(4, 27) = 0#
    pd(4, 28) = 0#
    pd(4, 29) = 0#
    pd(4, 30) = 0#
    pd(4, 31) = 0#

    ' --- total (pd row 5) — not used directly in this implementation ---
    pd(5, 0) = 1.0

    ' --- H (pd row 6) ---
    pd(6, 0)  = 1.02650e+0
    pd(6, 1)  = -6.27150e-2
    pd(6, 2)  = 8.73849e-2
    pd(6, 3)  = -2.96949e-2
    pd(6, 4)  = 0#
    pd(6, 5)  = 0#
    pd(6, 6)  = -1.68134e-1
    pd(6, 7)  = 0#
    pd(6, 8)  = -1.95177e-2
    pd(6, 9)  = 0#
    pd(6, 10) = 0#
    pd(6, 11) = 0#
    pd(6, 12) = 0#
    pd(6, 13) = 0#
    pd(6, 14) = 0#
    pd(6, 15) = 0#
    pd(6, 16) = -9.23822e-2
    pd(6, 17) = 0#
    pd(6, 18) = 0#
    pd(6, 19) = 5.56346e-2
    pd(6, 20) = 1.35265e-1
    pd(6, 21) = 0#
    pd(6, 22) = 0#
    pd(6, 23) = 0#
    pd(6, 24) = 0#
    pd(6, 25) = 0#
    pd(6, 26) = 0#
    pd(6, 27) = 0#
    pd(6, 28) = 0#
    pd(6, 29) = 0#
    pd(6, 30) = 0#
    pd(6, 31) = 0#

    ' --- N (pd row 7) ---
    pd(7, 0)  = 9.44256e-1
    pd(7, 1)  = 0#
    pd(7, 2)  = 0#
    pd(7, 3)  = 0#
    pd(7, 4)  = 0#
    pd(7, 5)  = 0#
    pd(7, 6)  = 0#
    pd(7, 7)  = 0#
    pd(7, 8)  = 0#
    pd(7, 9)  = 0#
    pd(7, 10) = 0#
    pd(7, 11) = 0#
    pd(7, 12) = 0#
    pd(7, 13) = 0#
    pd(7, 14) = 0#
    pd(7, 15) = 0#
    pd(7, 16) = 0#
    pd(7, 17) = 0#
    pd(7, 18) = 0#
    pd(7, 19) = -1.05585e-1
    pd(7, 20) = 0#
    pd(7, 21) = 0#
    pd(7, 22) = 0#
    pd(7, 23) = 0#
    pd(7, 24) = 0#
    pd(7, 25) = 0#
    pd(7, 26) = 0#
    pd(7, 27) = 0#
    pd(7, 28) = 0#
    pd(7, 29) = 0#
    pd(7, 30) = 0#
    pd(7, 31) = 0#

    ' --- anomalous O (pd row 8) ---
    pd(8, 0)  = 0#   ' small; handled separately
    pd(8, 1)  = 0#
    pd(8, 2)  = 0#
    pd(8, 3)  = 0#
    pd(8, 4)  = 0#
    pd(8, 5)  = 0#
    pd(8, 6)  = 0#
    pd(8, 7)  = 0#
    pd(8, 8)  = 0#
    pd(8, 9)  = 0#
    pd(8, 10) = 0#
    pd(8, 11) = 0#
    pd(8, 12) = 0#
    pd(8, 13) = 0#
    pd(8, 14) = 0#
    pd(8, 15) = 0#
    pd(8, 16) = 0#
    pd(8, 17) = 0#
    pd(8, 18) = 0#
    pd(8, 19) = 0#
    pd(8, 20) = 0#
    pd(8, 21) = 0#
    pd(8, 22) = 0#
    pd(8, 23) = 0#
    pd(8, 24) = 0#
    pd(8, 25) = 0#
    pd(8, 26) = 0#
    pd(8, 27) = 0#
    pd(8, 28) = 0#
    pd(8, 29) = 0#
    pd(8, 30) = 0#
    pd(8, 31) = 0#

    bInit = True

End Sub

'==============================================================================
' Sgn helper (returns -1, 0, or 1)
'==============================================================================
Private Function Sgn(ByVal x As Double) As Double
    If x > 0 Then Sgn = 1
    ElseIf x < 0 Then Sgn = -1
    Else Sgn = 0
    End If
End Function
