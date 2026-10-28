Attribute VB_Name = "modMain"
'==============================================================================
' modMain - buttons of the Pricer sheet
'
' Pricing of the vanilla structures (incl. the average rate option and the seagull) is done by live worksheet formulas (Schedule
' sheet, Pricer table, Scenarios grid). VBA only does what a formula cannot:
'   SolveZeroCost     solves the zero-cost strikes (collar give-up, seagull give-up,
'                     participating forward, TARF) and writes them into the strike cells
'   RunMonteCarlo     TARF mid value at the current strike + outcome statistics
'   PriceAll          both of the above
'   SetDefaultStrike  protection strike = budget if out-of-the-money, else fwd +/- 2.5%
'==============================================================================
Option Explicit

Private dMinY As Double, dMaxY As Double

Private Sub AppFast(ByVal fast As Boolean)
    On Error Resume Next
    Application.ScreenUpdating = Not fast
    If fast Then
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.Calculate
    End If
    On Error GoTo 0
End Sub

Private Sub SetStatus(ByVal msg As String)
    NamedCell("Status").Value = msg
    On Error Resume Next
    Application.StatusBar = msg
    On Error GoTo 0
End Sub

Private Function Stamp() As String
    Stamp = Format(Now, "yyyy-mm-dd hh:mm")
End Function

Public Sub SetDefaultStrike()
    LoadMarket
    SetupDeal
    If gPhi * (gBudget - gFbar) > 0.002 * gFbar Then
        NamedCell("Protect").Value = Round(gBudget, 4)
    Else
        NamedCell("Protect").Value = Round(gFbar * (1# + 0.025 * gPhi), 4)
    End If
    ' seagull: protection stops 5 big figures beyond the protection strike
    NamedCell("SgCap").Value = Round(CDbl(NamedValue("Protect")) + 0.05 * gPhi, 4)
End Sub

Public Sub PriceAll()
    Dim t0 As Double, m1 As String, m2 As String
    t0 = Timer
    m1 = DoSolve()
    m2 = DoMonteCarlo()
    SetStatus "Priced " & Stamp() & " | " & m1 & " | " & m2 & " | " & Format(Timer - t0, "0.0") & " s"
End Sub

Public Sub SolveZeroCost()
    Dim t0 As Double, m As String
    t0 = Timer
    m = DoSolve()
    SetStatus "Strikes solved " & Stamp() & " | " & m & " | " & Format(Timer - t0, "0.0") & _
              " s | Monte Carlo statistics not refreshed"
End Sub

Public Sub RunMonteCarlo()
    Dim t0 As Double, m As String
    t0 = Timer
    m = DoMonteCarlo()
    SetStatus "Monte Carlo " & Stamp() & " | " & m & " | " & Format(Timer - t0, "0.0") & " s"
End Sub

' Zero-cost strikes. Collar and participating forward in closed form (bisection on the
' Garman-Kohlhagen strip); TARF by Monte Carlo on risk-neutral paths.
Private Function DoSolve() As String
    Dim ok As Boolean, msg As String, k As Double, S() As Double
    On Error GoTo Fail
    AppFast True
    SetStatus "Solving zero-cost strikes..."
    LoadMarket
    SetupDeal
    k = SolveCollarGiveup(ok)
    If ok Then NamedCell("Giveup").Value = Round(k, 6) Else msg = msg & "no zero-cost collar; "
    k = SolveSeagullGiveup(ok)
    If ok Then NamedCell("SgGiveup").Value = Round(k, 6) Else msg = msg & "no zero-cost seagull; "
    k = SolvePFStrike(ok)
    If ok Then NamedCell("PFStrike").Value = Round(k, 6) Else msg = msg & "no zero-cost participating fwd; "
    SimulatePaths S, CLng(NamedValue("NPaths")), CLng(NamedValue("Seed")), True
    k = SolveTarfStrike(S, ok)
    If ok Then
        NamedCell("TarfStrike").Value = Round(k, 6)
        gTarfK = Round(k, 6)
        NamedCell("TarfPV").Value = TarfPV(S, gTarfK)
    Else
        msg = msg & "no zero-cost TARF; "
    End If
    AppFast False
    DoSolve = IIf(Len(msg) = 0, "zero-cost strikes OK", msg)
    Exit Function
Fail:
    AppFast False
    DoSolve = "Error in solver: " & Err.Description
End Function

' Achieved average rate (premium included) of strategy `code` on path p of S
Private Function PathRate(S() As Double, ByVal p As Long, ByVal code As Long, ByVal premFV As Double, _
                          overHedged As Boolean) As Double
    Dim i As Long, v As Double, ov As Double, c As Double, ko As Long, nl As Long, mo As Double, avg As Double
    overHedged = False
    For i = 1 To gNFix
        v = v + gPhi * gN / S(p, i)
    Next i
    If code = 8 Then
        v = v + TarfPath(S, p, gTarfK, c, ko, nl, mo)
        overHedged = (mo > 0)
    ElseIf code = 4 Then
        ' average rate option: one payoff on the average of the fixings, paid at the end
        avg = 0
        For i = 1 To gNFix
            avg = avg + S(p, i) / gNFix
        Next i
        If gPhi * (avg - gProtect) > 0 Then
            v = v + gN * gNFix / gProtect * gPhi * (avg - gProtect) / S(p, gNFix)
        End If
    Else
        For i = 1 To gNFix
            v = v + LegGain(code, S(p, i), i, ov)
        Next i
    End If
    v = v - premFV
    PathRate = gN * gNFix / Abs(v)
End Function

' TARF mid value at the strike in the sheet + statistics of every structure
Private Function DoMonteCarlo() As String
    Dim nP As Long, seed As Long, rnDrift As Boolean
    Dim S() As Double
    Dim code As Long, p As Long, j As Long
    Dim rates() As Double, oh As Boolean, nOver As Long, nWorse As Long, sumR As Double
    Dim premFV As Double, tAvg As Double
    Dim res(1 To 8, 1 To 5) As Variant
    Dim dist() As Variant, q As Double
    Dim c As Double, ko As Long, nl As Long, mo As Double
    Dim nKO As Long, sumLife As Double, nAnyLev As Long, sumLev As Double
    Dim qBad As Double, qGood As Double

    On Error GoTo Fail
    AppFast True
    LoadMarket
    SetupDeal
    nP = CLng(NamedValue("NPaths"))
    If nP < 1000 Then nP = 1000
    seed = CLng(NamedValue("Seed"))
    rnDrift = (InStr(1, CStr(NamedValue("DriftView")), "Forward", vbTextCompare) > 0)
    gPremium = StripOptionEUR(gProtect, ProtectIsCall(), gN)

    ' TARF mid value at the current strike (risk-neutral paths, same numbers as the solver)
    SetStatus "Monte Carlo: TARF value..."
    SimulatePaths S, nP, seed, True
    NamedCell("TarfPV").Value = TarfPV(S, gTarfK)
    Erase S

    ' statistics on fresh paths
    SetStatus "Monte Carlo: outcome statistics..."
    SimulatePaths S, nP, seed + 1, rnDrift
    nP = UBound(S, 1)
    tAvg = (gT(1) + gT(gNFix)) / 2#
    If gPhi > 0 Then qBad = 0.95: qGood = 0.05 Else qBad = 0.05: qGood = 0.95
    ReDim dist(1 To 99, 1 To N_STRAT + 1)
    For j = 1 To 99
        dist(j, 1) = j / 100#
    Next j
    dMinY = 1E+30: dMaxY = -1E+30
    For code = 1 To N_STRAT
        premFV = 0
        If code = 3 Then premFV = gPremium / DFFor(tAvg)
        If code = 4 Then premFV = CDbl(NamedValue("ApoPremium")) / DFFor(tAvg)
        ReDim rates(1 To nP)
        nOver = 0: nWorse = 0: sumR = 0
        For p = 1 To nP
            rates(p) = PathRate(S, p, code, premFV, oh)
            sumR = sumR + rates(p)
            If oh Then nOver = nOver + 1
            ' 1e-9 tolerance: a rate exactly at the budget (collar floor) is not a miss
            If gPhi > 0 Then
                If rates(p) > gBudget + 0.000000001 Then nWorse = nWorse + 1
            Else
                If rates(p) < gBudget - 0.000000001 Then nWorse = nWorse + 1
            End If
        Next p
        QuickSort rates, 1, nP
        res(code, 1) = sumR / nP
        res(code, 2) = PercentileSorted(rates, qBad)
        res(code, 3) = PercentileSorted(rates, qGood)
        res(code, 4) = nWorse / nP
        res(code, 5) = nOver / nP
        For j = 1 To 99
            ' quantile function, ordered from the client's best to worst outcome
            If gPhi > 0 Then q = j / 100# Else q = 1# - j / 100#
            dist(j, code + 1) = PercentileSorted(rates, q)
            If dist(j, code + 1) < dMinY Then dMinY = dist(j, code + 1)
            If dist(j, code + 1) > dMaxY Then dMaxY = dist(j, code + 1)
        Next j
    Next code

    ' TARF anatomy
    For p = 1 To nP
        TarfPath S, p, gTarfK, c, ko, nl, mo
        If ko <= gNFix Then nKO = nKO + 1: sumLife = sumLife + ko Else sumLife = sumLife + gNFix
        If nl > 0 Then nAnyLev = nAnyLev + 1
        sumLev = sumLev + nl
    Next p
    Erase S

    NamedCell("McTop").Resize(N_STRAT, 5).Value = res
    With NamedCell("TarfTop")
        .Offset(0, 0).Value = nKO / nP
        .Offset(1, 0).Value = sumLife / nP
        .Offset(2, 0).Value = nAnyLev / nP
        .Offset(3, 0).Value = sumLev / nP
    End With
    NamedCell("DistTop").Resize(99, N_STRAT + 1).Value = dist
    AppFast False
    RescaleCharts
    DoMonteCarlo = nP & " paths"
    Exit Function
Fail:
    AppFast False
    DoMonteCarlo = "Error in Monte Carlo: " & Err.Description
End Function

' Axis ranges follow the data (Excel would otherwise start the rate axis at zero)
Public Sub RescaleCharts()
    Dim rng As Range, cell As Range, lo As Double, hi As Double, xlo As Double, xhi As Double
    On Error Resume Next
    lo = 1E+30: hi = -1E+30
    Set rng = NamedCell("ScenTop").Offset(0, 2).Resize(61, N_STRAT)
    For Each cell In rng.Cells
        If IsNumeric(cell.Value) Then
            If cell.Value < lo Then lo = cell.Value
            If cell.Value > hi Then hi = cell.Value
        End If
    Next cell
    xlo = NamedCell("ScenTop").Offset(0, 1).Value
    xhi = NamedCell("ScenTop").Offset(60, 1).Value
    With ThisWorkbook.Worksheets("Pricer").ChartObjects(1).Chart
        .Axes(1).MinimumScale = Int(xlo * 20#) / 20#
        .Axes(1).MaximumScale = -Int(-xhi * 20#) / 20#
        .Axes(2).MinimumScale = Int(lo * 20#) / 20#
        .Axes(2).MaximumScale = -Int(-hi * 20#) / 20#
    End With
    If dMaxY > dMinY Then
        With ThisWorkbook.Worksheets("Distribution").ChartObjects(1).Chart
            .Axes(2).MinimumScale = Int(dMinY * 50#) / 50#
            .Axes(2).MaximumScale = -Int(-dMaxY * 50#) / 50#
        End With
    End If
    On Error GoTo 0
End Sub
