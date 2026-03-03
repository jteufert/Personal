Attribute VB_Name = "NRLMSISE00_PythonBridge"
'==============================================================================
' NRLMSISE-00 Python Bridge (optional, higher-accuracy alternative)
'
' If Python is installed with the msise00 or nrlmsise00 package, this module
' calls nrlmsise_calc.py via Shell and reads the result from a temp file.
' This gives you verified NRLMSISE-00 results instead of the VBA implementation.
'
' Usage (in a cell):
'   =ATMDENSITY_PY(400, 150)        -> density at 400 km, F10.7 = 150
'
' To install msise00:
'   pip install msise00
'
' Place nrlmsise_calc.py in the same folder as this workbook, or set
' PYTHON_SCRIPT_PATH below to its absolute path.
'==============================================================================
Option Explicit

' ---- CONFIGURE THIS ----
Private Const PYTHON_EXE As String = "python"   ' or "python3", or full path like "C:\Python312\python.exe"
' ---- END CONFIG --------

'------------------------------------------------------------------------------
' ATMDENSITY_PY — calls Python nrlmsise_calc.py, returns kg/m^3
'------------------------------------------------------------------------------
Public Function ATMDENSITY_PY(alt_km As Double, f107 As Double, _
    Optional f107a As Double = 0, _
    Optional Ap    As Double = 4, _
    Optional doy   As Long   = 172, _
    Optional lat   As Double = 45, _
    Optional lon   As Double = 0, _
    Optional lst   As Double = 12) As Double

    If f107a = 0 Then f107a = f107

    Dim scriptPath As String
    scriptPath = GetScriptPath()
    If scriptPath = "" Then
        ATMDENSITY_PY = CVErr(xlErrNA)
        Exit Function
    End If

    Dim tmpFile As String
    tmpFile = Environ("TEMP") & "\nrlmsise_result_" & Format(Now, "hhmmss") & ".txt"

    ' Build command: python nrlmsise_calc.py alt f107 f107a ap doy lat lon lst > tmpFile
    Dim cmd As String
    cmd = PYTHON_EXE & " """ & scriptPath & """ " & _
          alt_km & " " & f107 & " " & f107a & " " & Ap & " " & _
          doy & " " & lat & " " & lon & " " & lst & _
          " > """ & tmpFile & """ 2>&1"

    ' Run synchronously via cmd /c
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim exitCode As Long
    exitCode = wsh.Run("cmd /c " & cmd, 0, True)   ' 0=hidden, True=wait

    ' Read result
    If Dir(tmpFile) = "" Then
        ATMDENSITY_PY = CVErr(xlErrNA)
        Exit Function
    End If

    Dim fNum As Integer: fNum = FreeFile
    Open tmpFile For Input As #fNum
    Dim line1 As String
    Line Input #fNum, line1
    Close #fNum
    Kill tmpFile

    line1 = Trim(line1)
    If Left(line1, 5) = "ERROR" Or Not IsNumeric(line1) Then
        ' Surface the error in the cell comment (optional)
        ATMDENSITY_PY = CVErr(xlErrNA)
    Else
        ATMDENSITY_PY = CDbl(line1)
    End If

End Function

'------------------------------------------------------------------------------
' GetScriptPath — find nrlmsise_calc.py next to workbook or add-in
'------------------------------------------------------------------------------
Private Function GetScriptPath() As String
    Dim candidate As String

    ' 1. Same folder as the active workbook
    If ActiveWorkbook.Path <> "" Then
        candidate = ActiveWorkbook.Path & "\nrlmsise_calc.py"
        If Dir(candidate) <> "" Then
            GetScriptPath = candidate
            Exit Function
        End If
    End If

    ' 2. Same folder as this add-in (ThisWorkbook)
    candidate = ThisWorkbook.Path & "\nrlmsise_calc.py"
    If Dir(candidate) <> "" Then
        GetScriptPath = candidate
        Exit Function
    End If

    GetScriptPath = ""
End Function
