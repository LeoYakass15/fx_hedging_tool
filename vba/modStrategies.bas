Attribute VB_Name = "modStrategies"
'==============================================================================
' modStrategies - hedging structures on a strip of monthly fixings.
'
' Every structure is a set of "conversion legs": an amount n of USD converted at a
' hedge rate R instead of the spot S. Home-currency (EUR) gain of a leg:
'       gain = phi * n * (1/R - 1/S)       phi = +1 exporter, -1 importer
' Unhedged value of a fixing = phi * N / S (EUR received, or minus EUR paid).
' Mid value to the client = sum_i DF_USD(t_i) / S0 * E_Q[ gain_i * S_i ].
'
' Strategy codes: 1 Unhedged, 2 Forward, 3 Vanilla option, 4 Zero-cost collar,
'                 5 Participating forward, 6 TARF
'==============================================================================
Option Explicit

Public Const N_STRAT As Long = 6

' Deal state (filled by SetupDeal)
Public gPhi As Double            ' +1 receivable (exporter), -1 payable (importer)
Public gN As Double              ' USD per fixing
Public gNFix As Long
Public gT() As Double            ' fixing times (years)
Public gF() As Double            ' outright forwards
Public gDFd() As Double          ' USD discount factors
Public gBudget As Double
Public gMarginEUR As Double      ' sales margin in EUR (zero-cost structures)
Public gFwdRate() As Double      ' forward strip rates incl. margin
Public gProtect As Double, gGiveup As Double
Public gPFStrike As Double, gPart As Double
Public gTarfK As Double, gTarfTarget As Double, gLev As Double, gExactKO As Boolean
Public gPremium As Double        ' option premium, EUR
Public gFbar As Double

Public Sub SetupDeal()
    Dim i As Long, marginBp As Double
    gPhi = IIf(InStr(1, CStr(NamedValue("Side")), "Export", vbTextCompare) > 0, 1#, -1#)
    gN = CDbl(NamedValue("Notional"))
    gNFix = CLng(NamedValue("NMonths"))
    If gNFix < 1 Then gNFix = 1
    If gNFix > 24 Then gNFix = 24
    gBudget = CDbl(NamedValue("Budget"))
    gPart = CDbl(NamedValue("Participation"))
    gTarfTarget = CDbl(NamedValue("TarfTarget")) / 10000#
    gLev = CDbl(NamedValue("Leverage"))
    gExactKO = (LCase(CStr(NamedValue("KOType"))) = "exact")
    marginBp = CDbl(NamedValue("MarginBp"))
    gProtect = CDbl(NamedValue("Protect"))
    gGiveup = CDbl(NamedValue("Giveup"))
    gPFStrike = CDbl(NamedValue("PFStrike"))
    gTarfK = CDbl(NamedValue("TarfStrike"))
    ReDim gT(1 To gNFix): ReDim gF(1 To gNFix): ReDim gDFd(1 To gNFix): ReDim gFwdRate(1 To gNFix)
    gFbar = 0
    For i = 1 To gNFix
        gT(i) = i / 12#
        gF(i) = Fwd(gT(i))
        gDFd(i) = DFDom(gT(i))
        gFbar = gFbar + gF(i) / gNFix
    Next i
    ' margin: bp of the total home-currency notional; on the forward, bp of the rate
    gMarginEUR = marginBp * 0.0001 * gN * gNFix / mSpot
    For i = 1 To gNFix
        gFwdRate(i) = gF(i) - gPhi * marginBp * 0.0001 * gFbar
    Next i
End Sub

Public Function ProtectIsCall() As Boolean
    ProtectIsCall = (gPhi > 0)
End Function

' EUR value of options on notionalUSD/K units of EUR, summed over the strip
Public Function StripOptionEUR(ByVal K As Double, ByVal isCall As Boolean, ByVal notionalUSD As Double) As Double
    Dim i As Long, s As Double
    For i = 1 To gNFix
        s = s + notionalUSD / K * PriceMkt(K, gT(i), isCall) / mSpot
    Next i
    StripOptionEUR = s
End Function

Public Function CollarPV(ByVal kProt As Double, ByVal kGive As Double) As Double
    CollarPV = StripOptionEUR(kProt, ProtectIsCall(), gN) - StripOptionEUR(kGive, Not ProtectIsCall(), gN)
End Function

Public Function PFPV(ByVal K As Double, ByVal part As Double) As Double
    PFPV = StripOptionEUR(K, ProtectIsCall(), gN) - StripOptionEUR(K, Not ProtectIsCall(), (1# - part) * gN)
End Function

' Bisection root finder on f(x) = target. which: 1 collar give-up, 2 PF strike
Private Function EvalStrike(ByVal which As Long, ByVal x As Double) As Double
    Select Case which
        Case 1: EvalStrike = CollarPV(gProtect, x) + gMarginEUR
        Case 2: EvalStrike = PFPV(x, gPart) + gMarginEUR
    End Select
End Function

Private Function Bisect(ByVal which As Long, ByVal lo As Double, ByVal hi As Double, ok As Boolean) As Double
    Dim flo As Double, fhi As Double, mid As Double, fm As Double, it As Long
    flo = EvalStrike(which, lo): fhi = EvalStrike(which, hi)
    ok = (flo * fhi <= 0)
    If Not ok Then Bisect = 0: Exit Function
    For it = 1 To 60
        mid = 0.5 * (lo + hi)
        fm = EvalStrike(which, mid)
        If fm * flo > 0 Then
            lo = mid: flo = fm
        Else
            hi = mid
        End If
        If hi - lo < 0.0000000001 Then Exit For
    Next it
    Bisect = 0.5 * (lo + hi)
End Function

Public Function SolveCollarGiveup(ok As Boolean) As Double
    Dim lo As Double, hi As Double
    If gPhi > 0 Then
        lo = 0.5 * gFbar
        hi = IIf(gProtect < gFbar, gProtect, gFbar) * 0.99999
    Else
        lo = IIf(gProtect > gFbar, gProtect, gFbar) * 1.00001
        hi = 2# * gFbar
    End If
    SolveCollarGiveup = Bisect(1, lo, hi, ok)
End Function

Public Function SolvePFStrike(ok As Boolean) As Double
    SolvePFStrike = Bisect(2, 0.7 * gFbar, 1.3 * gFbar, ok)
End Function

' Hedge gain on one fixing for the non path-dependent structures
Public Function LegGain(ByVal code As Long, ByVal S As Double, ByVal i As Long, overN As Double) As Double
    Dim g As Double, ph As Double
    ph = gPhi
    overN = 0
    Select Case code
        Case 1
            g = 0
        Case 2
            g = ph * gN * (1# / gFwdRate(i) - 1# / S)
        Case 3
            If ph * (S - gProtect) > 0 Then g = ph * gN * (1# / gProtect - 1# / S)
        Case 4
            If ph * (S - gProtect) > 0 Then g = g + ph * gN * (1# / gProtect - 1# / S)
            If ph * (gGiveup - S) > 0 Then g = g + ph * gN * (1# / gGiveup - 1# / S)
        Case 5
            If ph * (S - gPFStrike) > 0 Then
                g = ph * gN * (1# / gPFStrike - 1# / S)
            Else
                g = ph * (1# - gPart) * gN * (1# / gPFStrike - 1# / S)
            End If
    End Select
    LegGain = g
End Function

' TARF on one path. S(p, 1..n). Returns total EUR gain; also discounted E[gain*S]
' contribution (pvContrib), knock-out fixing (koIdx = n+1 if alive), leveraged count.
Public Function TarfPath(S() As Double, ByVal p As Long, ByVal K As Double, _
                         pvContrib As Double, koIdx As Long, nLev As Long, maxOver As Double) As Double
    Dim i As Long, s_ As Double, intr As Double, acc As Double, nleg As Double
    Dim scl As Double, g As Double, tot As Double, alive As Boolean
    alive = True: acc = 0: tot = 0: pvContrib = 0: nLev = 0: koIdx = gNFix + 1: maxOver = 0
    For i = 1 To gNFix
        If Not alive Then Exit For
        s_ = S(p, i)
        intr = gPhi * (s_ - K)
        If intr > 0 Then
            scl = 1#
            If gExactKO Then
                If (gTarfTarget - acc) < intr Then scl = (gTarfTarget - acc) / intr
                If scl < 0 Then scl = 0
            End If
            nleg = gN * scl
            acc = acc + intr * scl
        Else
            nleg = gLev * gN
            nLev = nLev + 1
            If nleg - gN > maxOver Then maxOver = nleg - gN
        End If
        g = gPhi * nleg * (1# / K - 1# / s_)
        tot = tot + g
        pvContrib = pvContrib + gDFd(i) / mSpot * g * s_
        If acc >= gTarfTarget - 0.00000000000001 Then
            alive = False
            koIdx = i
        End If
    Next i
    TarfPath = tot
End Function
