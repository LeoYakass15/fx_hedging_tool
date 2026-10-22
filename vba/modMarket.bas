Attribute VB_Name = "modMarket"
'==============================================================================
' modMarket - EURUSD market: curves, forwards, ATM term structure and smile,
'             Garman-Kohlhagen pricing. Also exposes worksheet functions (UDFs).
'
' Convention: EURUSD = USD per 1 EUR. Domestic (quote) = USD, foreign (base) = EUR.
' Market inputs live on the "Market" sheet (named range MktTable):
'   years | USD zero rate % | EUR zero rate % | ATM vol % | 25d RR % | 25d BF %
' Spot and the vol shift come from the Pricer sheet (names Spot, VolShift).
' The worksheet functions (FX_FORWARD, FX_DFDOM, FX_DFFOR, FX_VOL...) feed the live
' formulas of the Schedule and Market sheets.
'==============================================================================
Option Explicit

Public mN As Long
Public mT() As Double, mRd() As Double, mRf() As Double
Public mAtm() As Double, mRR() As Double, mBF() As Double
Public mSpot As Double

Public Function NamedValue(ByVal nm As String) As Variant
    NamedValue = ThisWorkbook.Names(nm).RefersToRange.Value
End Function

Public Function NamedCell(ByVal nm As String) As Range
    Set NamedCell = ThisWorkbook.Names(nm).RefersToRange
End Function

' Reads the market table into module arrays. spotOverride/volShift < -998 = use sheet.
Public Sub LoadMarket(Optional ByVal spotOverride As Double = -999#, _
                      Optional ByVal volShiftPts As Double = -999#)
    Dim tbl As Range, r As Long, n As Long
    Set tbl = NamedCell("MktTable")
    n = 0
    For r = 1 To tbl.Rows.Count
        If Len(CStr(tbl.Cells(r, 1).Value)) > 0 Then n = n + 1
    Next r
    mN = n
    ReDim mT(1 To n): ReDim mRd(1 To n): ReDim mRf(1 To n)
    ReDim mAtm(1 To n): ReDim mRR(1 To n): ReDim mBF(1 To n)
    If volShiftPts < -998# Then volShiftPts = CDbl(NamedValue("VolShift"))
    n = 0
    For r = 1 To tbl.Rows.Count
        If Len(CStr(tbl.Cells(r, 1).Value)) > 0 Then
            n = n + 1
            mT(n) = CDbl(tbl.Cells(r, 1).Value)
            mRd(n) = CDbl(tbl.Cells(r, 2).Value) / 100#
            mRf(n) = CDbl(tbl.Cells(r, 3).Value) / 100#
            mAtm(n) = (CDbl(tbl.Cells(r, 4).Value) + volShiftPts) / 100#
            mRR(n) = CDbl(tbl.Cells(r, 5).Value) / 100#
            mBF(n) = CDbl(tbl.Cells(r, 6).Value) / 100#
        End If
    Next r
    If spotOverride < -998# Then
        mSpot = CDbl(NamedValue("Spot"))
    Else
        mSpot = spotOverride
    End If
End Sub

Public Function DFDom(ByVal T As Double) As Double
    DFDom = Exp(-Interp(T, mT, mRd) * T)
End Function

Public Function DFFor(ByVal T As Double) As Double
    DFFor = Exp(-Interp(T, mT, mRf) * T)
End Function

' Outright forward by covered interest parity
Public Function Fwd(ByVal T As Double) As Double
    Fwd = mSpot * DFFor(T) / DFDom(T)
End Function

' ATM vol, linear interpolation in total variance, flat vol before the first pillar
Public Function AtmVol(ByVal T As Double) As Double
    Dim w() As Double, i As Long
    If T < 0.000001 Then T = 0.000001
    If T <= mT(1) Then AtmVol = mAtm(1): Exit Function
    ReDim w(1 To mN)
    For i = 1 To mN
        w(i) = mAtm(i) * mAtm(i) * mT(i)
    Next i
    If T >= mT(mN) Then AtmVol = mAtm(mN): Exit Function
    AtmVol = Sqr(Interp(T, mT, w) / T)
End Function

' Smile: quadratic in forward call delta x through 3 market points
'   x=0.25 -> ATM + BF + RR/2 ; x=0.5 -> ATM ; x=0.75 -> ATM + BF - RR/2
' Fixed-point iteration because delta depends on the vol.
Public Function SmileVol(ByVal K As Double, ByVal T As Double) As Double
    Dim atm As Double, rr As Double, bf As Double, F As Double
    Dim b As Double, c As Double, sig As Double, nw As Double, d1 As Double, x As Double
    Dim it As Long
    atm = AtmVol(T)
    rr = Interp(T, mT, mRR)
    bf = Interp(T, mT, mBF)
    F = Fwd(T)
    b = -2# * rr: c = 16# * bf
    sig = atm
    For it = 1 To 60
        d1 = (Log(F / K) + 0.5 * sig * sig * T) / (sig * Sqr(T))
        x = NormCDF(d1)
        If x < 0.02 Then x = 0.02
        If x > 0.98 Then x = 0.98
        nw = atm + b * (x - 0.5) + c * (x - 0.5) ^ 2
        If nw < 0.01 Then nw = 0.01
        If Abs(nw - sig) < 0.0000000001 Then sig = nw: Exit For
        sig = 0.5 * sig + 0.5 * nw
    Next it
    SmileVol = sig
End Function

' Garman-Kohlhagen price in quote ccy (USD) per 1 unit of base (EUR)
Public Function GKPrice(ByVal S As Double, ByVal K As Double, ByVal T As Double, _
                        ByVal rd As Double, ByVal rf As Double, ByVal vol As Double, _
                        ByVal isCall As Boolean) As Double
    Dim d1 As Double, d2 As Double, vs As Double
    If T < 0.000000000001 Then T = 0.000000000001
    vs = vol * Sqr(T)
    d1 = (Log(S / K) + (rd - rf + 0.5 * vol * vol) * T) / vs
    d2 = d1 - vs
    If isCall Then
        GKPrice = S * Exp(-rf * T) * NormCDF(d1) - K * Exp(-rd * T) * NormCDF(d2)
    Else
        GKPrice = K * Exp(-rd * T) * NormCDF(-d2) - S * Exp(-rf * T) * NormCDF(-d1)
    End If
End Function

' Price with market curves and smile
Public Function PriceMkt(ByVal K As Double, ByVal T As Double, ByVal isCall As Boolean) As Double
    Dim rd As Double, rf As Double
    rd = -Log(DFDom(T)) / T
    rf = -Log(DFFor(T)) / T
    PriceMkt = GKPrice(mSpot, K, T, rd, rf, SmileVol(K, T), isCall)
End Function

'------------------------------------------------------------------ worksheet UDFs
Public Function FX_FORWARD(ByVal T As Double) As Double
    Application.Volatile
    LoadMarket
    FX_FORWARD = Fwd(T)
End Function

Public Function FX_POINTS(ByVal T As Double) As Double
    Application.Volatile
    LoadMarket
    FX_POINTS = (Fwd(T) - mSpot) * 10000#
End Function

Public Function FX_VOL(ByVal K As Double, ByVal T As Double) As Double
    Application.Volatile
    LoadMarket
    FX_VOL = SmileVol(K, T)
End Function

Public Function FX_OPTION(ByVal K As Double, ByVal T As Double, ByVal callPut As String) As Double
    Application.Volatile
    LoadMarket
    FX_OPTION = PriceMkt(K, T, UCase(Left(callPut, 1)) = "C")
End Function

' Average outright forward over the first n monthly fixings
Public Function FX_AVGFWD(ByVal nMonths As Long) As Double
    Dim i As Long, s As Double
    Application.Volatile
    LoadMarket
    For i = 1 To nMonths
        s = s + Fwd(i / 12#)
    Next i
    FX_AVGFWD = s / nMonths
End Function

Public Function FX_DFDOM(ByVal T As Double) As Double
    Application.Volatile
    LoadMarket
    FX_DFDOM = DFDom(T)
End Function

Public Function FX_DFFOR(ByVal T As Double) As Double
    Application.Volatile
    LoadMarket
    FX_DFFOR = DFFor(T)
End Function

Public Function FX_ATMVOL(ByVal T As Double) As Double
    Application.Volatile
    LoadMarket
    FX_ATMVOL = AtmVol(T)
End Function
