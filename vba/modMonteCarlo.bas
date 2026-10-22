Attribute VB_Name = "modMonteCarlo"
'==============================================================================
' modMonteCarlo - log-normal spot paths on the fixing dates, ATM vol term structure
'   drift "risk-neutral": E[S(t)] = F(t)  (pricing, TARF zero-cost strike)
'   drift "spot unchanged": E[S(t)] = S0  (desk view used for statistics)
' Antithetic variates, MRG32k3a generator (same numbers in every Excel version).
'==============================================================================
Option Explicit

Public Sub SimulatePaths(S() As Double, ByVal nPaths As Long, ByVal seed As Long, ByVal riskNeutral As Boolean)
    Dim half As Long, p As Long, i As Long
    Dim tv() As Double, dv() As Double, ctr() As Double, z() As Double
    Dim x As Double, xa As Double
    If nPaths Mod 2 = 1 Then nPaths = nPaths + 1
    half = nPaths \ 2
    ReDim S(1 To nPaths, 1 To gNFix)
    ReDim tv(0 To gNFix): ReDim dv(1 To gNFix): ReDim ctr(1 To gNFix): ReDim z(1 To gNFix)
    tv(0) = 0
    For i = 1 To gNFix
        tv(i) = AtmVol(gT(i)) ^ 2 * gT(i)
        dv(i) = tv(i) - tv(i - 1)
        If dv(i) < 0.000000000001 Then dv(i) = 0.000000000001
        If riskNeutral Then ctr(i) = gF(i) Else ctr(i) = mSpot
    Next i
    RngSeed seed
    For p = 1 To half
        x = 0: xa = 0
        For i = 1 To gNFix
            z(i) = RngNormal() * Sqr(dv(i))
            x = x + z(i)
            xa = xa - z(i)
            S(p, i) = ctr(i) * Exp(x - 0.5 * tv(i))
            S(p + half, i) = ctr(i) * Exp(xa - 0.5 * tv(i))
        Next i
    Next p
End Sub

' Mid value of a TARF with strike K on given risk-neutral paths (EUR)
Public Function TarfPV(S() As Double, ByVal K As Double) As Double
    Dim p As Long, nP As Long, pv As Double, c As Double, ko As Long, nl As Long, mo As Double
    nP = UBound(S, 1)
    For p = 1 To nP
        TarfPath S, p, K, c, ko, nl, mo
        pv = pv + c
    Next p
    TarfPV = pv / nP
End Function

' Zero-cost TARF strike by bisection, common random numbers
Public Function SolveTarfStrike(S() As Double, ok As Boolean) As Double
    Dim lo As Double, hi As Double, flo As Double, fhi As Double, mid As Double, fm As Double
    Dim it As Long
    lo = 0.7 * gFbar: hi = 1.3 * gFbar
    flo = TarfPV(S, lo) + gMarginEUR
    fhi = TarfPV(S, hi) + gMarginEUR
    ok = (flo * fhi <= 0)
    If Not ok Then SolveTarfStrike = 0: Exit Function
    For it = 1 To 40
        mid = 0.5 * (lo + hi)
        fm = TarfPV(S, mid) + gMarginEUR
        If fm * flo > 0 Then
            lo = mid: flo = fm
        Else
            hi = mid
        End If
        If hi - lo < 0.0000001 Then Exit For
    Next it
    SolveTarfStrike = 0.5 * (lo + hi)
End Function
