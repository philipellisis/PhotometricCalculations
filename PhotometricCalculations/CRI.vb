Public Class CRICoordinates
    Public upp, vpp, WS, US, VS, c, d, u, v, X, Y, Z, DeltaE, Ri, Y0 As Double
    Private denomator, dDen, cDen As Double
    Private i As Integer
    Public Function calculateMainTEST(spd() As Double) As Boolean
        Y0 = 0
        X = 0
        Y = 0
        Z = 0
        For Me.i = 0 To 400
            Y0 += spd(i) * CIE2DegreeFunctions.ybar(i)
            X += spd(i) * CIE2DegreeFunctions.xbar(i)
            Y += spd(i) * CIE2DegreeFunctions.ybar(i)
            Z += spd(i) * CIE2DegreeFunctions.zbar(i)
        Next


        If Y0 = 0 Then
            Return 0
        End If

        denomator = 100 / Y0
        X = X * denomator
        Y = Y * denomator
        Z = Z * denomator
        denomator = (X + 15 * Y + 3 * Z)
        If denomator = 0 Then
            Return 0
        End If
        denomator = 1 / denomator
        u = 4 * X * denomator
        v = 6 * Y * denomator
        If v = 0 Then
            Return 0
        End If
        denomator = 1 / v
        d = (1.708 * v + 0.404 - 1.481 * u) * denomator
        c = (4 - u - 10 * v) * denomator
        Return 1
    End Function
    Public Function calculateMainREF(cri As CRI) As Boolean
        Y0 = 0
        X = 0
        Y = 0
        Z = 0
        For Me.i = 0 To 400
            Y0 += cri.ref(i) * CIE2DegreeFunctions.ybar(i)
            X += cri.ref(i) * CIE2DegreeFunctions.xbar(i)
            Y += cri.ref(i) * CIE2DegreeFunctions.ybar(i)
            Z += cri.ref(i) * CIE2DegreeFunctions.zbar(i)
        Next


        If Y0 = 0 Then
            Return 0
        End If
        denomator = 100 / Y0
        X = X * denomator
        Y = Y * denomator
        Z = Z * denomator
        denomator = (X + 15 * Y + 3 * Z)
        If denomator = 0 Then
            Return 0
        End If
        denomator = 1 / denomator
        u = 4 * X * denomator
        v = 6 * Y * denomator
        If v = 0 Then
            Return 0
        End If
        denomator = 1 / v
        d = (1.708 * v + 0.404 - 1.481 * u) * denomator
        c = (4 - u - 10 * v) * denomator
        Return 1
    End Function
    Public Function calculateREF(cri As CRI, ByRef color() As Double, RefCoordinates As CRICoordinates) As Integer
        Y0 = 0
        X = 0
        Y = 0
        Z = 0
        For Me.i = 0 To 400
            Y0 += cri.ref(i) * CIE2DegreeFunctions.ybar(i)
            X += cri.ref(i) * CIE2DegreeFunctions.xbar(i) * color(i)
            Y += cri.ref(i) * CIE2DegreeFunctions.ybar(i) * color(i)
            Z += cri.ref(i) * CIE2DegreeFunctions.zbar(i) * color(i)
        Next
        If Y0 <= 0 Or Y <= 0 Then
            Return 0
        End If
        denomator = 100 / Y0
        X = X * denomator
        Y = Y * denomator
        Z = Z * denomator
        denomator = (X + 15 * Y + 3 * Z)
        If denomator = 0 Then
            Return 0
        End If
        denomator = 1 / denomator
        u = 4 * X * denomator
        v = 6 * Y * denomator
        If v = 0 Then
            Return 0
        End If
        denomator = 1 / v
        d = (1.708 * v + 0.404 - 1.481 * u) * denomator
        c = (4 - u - 10 * v) * denomator
        WS = 25 * ((Y) ^ (1 / 3)) - 17
        US = 13 * WS * (u - RefCoordinates.u)
        VS = 13 * WS * (v - RefCoordinates.v)
        Return 1
    End Function
    Public Function calculateTEST(spd() As Double, ByRef color() As Double, RefCoordinates As CRICoordinates, TestCoordinates As CRICoordinates)
        Y0 = 0
        X = 0
        Y = 0
        Z = 0
        For Me.i = 0 To 400
            Y0 += spd(i) * CIE2DegreeFunctions.ybar(i)
            X += spd(i) * CIE2DegreeFunctions.xbar(i) * color(i)
            Y += spd(i) * CIE2DegreeFunctions.ybar(i) * color(i)
            Z += spd(i) * CIE2DegreeFunctions.zbar(i) * color(i)
        Next
        If Y0 <= 0 Or Y <= 0 Then
            Return 0
        End If
        denomator = 100 / Y0
        X = X * denomator
        Y = Y * denomator
        Z = Z * denomator
        denomator = (X + 15 * Y + 3 * Z)
        If denomator = 0 Then
            Return 0
        End If
        denomator = 1 / denomator
        u = 4 * X * denomator
        v = 6 * Y * denomator
        If v = 0 Then
            Return 0
        End If
        denomator = 1 / v
        d = (1.708 * v + 0.404 - 1.481 * u) * denomator
        c = (4 - u - 10 * v) * denomator
        WS = 25 * ((Y) ^ (1 / 3)) - 17
        dDen = (RefCoordinates.d / TestCoordinates.d) * d
        cDen = (RefCoordinates.c / TestCoordinates.c) * c
        denomator = 1 / (16.518 + 1.481 * cDen - dDen)
        upp = (10.872 + 0.404 * cDen - 4 * dDen) * denomator
        vpp = 5.52 * denomator

        US = 13 * WS * (upp - RefCoordinates.u)
        VS = 13 * WS * (vpp - RefCoordinates.v)
        Return 1
    End Function
    Public Sub calculateDeltaE(RefCoordinates As CRICoordinates)
        DeltaE = ((RefCoordinates.US - US) ^ 2 + (RefCoordinates.VS - VS) ^ 2 + (RefCoordinates.WS - WS) ^ 2) ^ 0.5
        Ri = 100 - 4.6 * DeltaE
    End Sub
End Class



Public Class CRI
    Private s(400) As Double
    Private plank(400) As Double
    Public ref(400) As Double
    'Public Y0 As Double
    Public CRI As Double
    Public R9 As Double
    Private MainRef As New CRICoordinates
    Private MainTest As New CRICoordinates
    Private CRIREF(8) As CRICoordinates
    Private CRITEST(8) As CRICoordinates
    Private i As Integer
    Private continueMeasure As Boolean
    Private ccx, ccy, m1, m2 As Double
    Public Sub New()
        For Me.i = 0 To 8
            CRIREF(i) = New CRICoordinates
            CRITEST(i) = New CRICoordinates
        Next
    End Sub

    ''' <summary>
    ''' Simple routine to calculate CRI. CRI calculations are simple set in the seperate publicly exposed parameters
    ''' </summary>
    ''' <param name="spd">Ensure that the wavelength is from 380-780 in 1nm increments</param>
    ''' <param name="cct">Precalculated CCT</param>
    Public Sub calculate(spd() As Double, cct As Double)
        continueMeasure = True
        If cct < 5000 Then
            For Me.i = 0 To 400
                'this also works, but takes a lot longer to calculate. The criRef is there to help with performance
                'ref(i) = (0.00000000000000037415 * (((i + 380) * 0.000000001) ^ -5)) / (2.71828182845905 ^ (0.014388 / ((i + 380) * 0.000000001 * cct)) - 1)
                ref(i) = CIE2DegreeFunctions.criRef(i) / (2.71828182845905 ^ (0.014388 / (CIE2DegreeFunctions.criRef2(i) * cct)) - 1)
            Next
        Else
            If cct < 7000 Then
                ccx = -4.607 * (10 ^ 9 / cct ^ 3) + 2.9678 * (10 ^ 6 / cct ^ 2) + 0.09911 * (10 ^ 3 / cct) + 0.244063
            Else
                ccx = -2.0064 * (10 ^ 9 / cct ^ 3) + 1.9018 * (10 ^ 6 / cct ^ 2) + 0.24748 * (10 ^ 3 / cct) + 0.23704
            End If
            ccy = -3 * ccx ^ 2 + 2.87 * ccx - 0.275
            m1 = (-1.3515 - 1.7703 * ccx + 5.9114 * ccy) / (0.0241 + 0.2562 * ccx - 0.7341 * ccy)
            m2 = (0.03 - 31.4424 * ccx + 30.0717 * ccy) / (0.0241 + 0.2562 * ccx - 0.7341 * ccy)
            For Me.i = 0 To 400
                ref(i) = CIE2DegreeFunctions.s0(i) + CIE2DegreeFunctions.s1(i) * m1 + CIE2DegreeFunctions.s2(i) * m2
            Next
        End If

        If continueMeasure Then continueMeasure = MainRef.calculateMainREF(Me)
        If continueMeasure Then continueMeasure = MainTest.calculateMainTEST(spd)

        If continueMeasure Then continueMeasure = CRIREF(0).calculateREF(Me, CIE2DegreeFunctions.p1, MainRef)
        If continueMeasure Then continueMeasure = CRIREF(1).calculateREF(Me, CIE2DegreeFunctions.p2, MainRef)
        If continueMeasure Then continueMeasure = CRIREF(2).calculateREF(Me, CIE2DegreeFunctions.p3, MainRef)
        If continueMeasure Then continueMeasure = CRIREF(3).calculateREF(Me, CIE2DegreeFunctions.p4, MainRef)
        If continueMeasure Then continueMeasure = CRIREF(4).calculateREF(Me, CIE2DegreeFunctions.p5, MainRef)
        If continueMeasure Then continueMeasure = CRIREF(5).calculateREF(Me, CIE2DegreeFunctions.p6, MainRef)
        If continueMeasure Then continueMeasure = CRIREF(6).calculateREF(Me, CIE2DegreeFunctions.p7, MainRef)
        If continueMeasure Then continueMeasure = CRIREF(7).calculateREF(Me, CIE2DegreeFunctions.p8, MainRef)
        If continueMeasure Then continueMeasure = CRIREF(8).calculateREF(Me, CIE2DegreeFunctions.p9, MainRef)

        If continueMeasure Then continueMeasure = CRITEST(0).calculateTEST(spd, CIE2DegreeFunctions.p1, MainRef, MainTest)
        If continueMeasure Then continueMeasure = CRITEST(1).calculateTEST(spd, CIE2DegreeFunctions.p2, MainRef, MainTest)
        If continueMeasure Then continueMeasure = CRITEST(2).calculateTEST(spd, CIE2DegreeFunctions.p3, MainRef, MainTest)
        If continueMeasure Then continueMeasure = CRITEST(3).calculateTEST(spd, CIE2DegreeFunctions.p4, MainRef, MainTest)
        If continueMeasure Then continueMeasure = CRITEST(4).calculateTEST(spd, CIE2DegreeFunctions.p5, MainRef, MainTest)
        If continueMeasure Then continueMeasure = CRITEST(5).calculateTEST(spd, CIE2DegreeFunctions.p6, MainRef, MainTest)
        If continueMeasure Then continueMeasure = CRITEST(6).calculateTEST(spd, CIE2DegreeFunctions.p7, MainRef, MainTest)
        If continueMeasure Then continueMeasure = CRITEST(7).calculateTEST(spd, CIE2DegreeFunctions.p8, MainRef, MainTest)
        If continueMeasure Then continueMeasure = CRITEST(8).calculateTEST(spd, CIE2DegreeFunctions.p9, MainRef, MainTest)

        If continueMeasure Then
            CRI = 0
            For Me.i = 0 To 8
                CRITEST(i).calculateDeltaE(CRIREF(i))
                CRI += CRITEST(i).Ri
            Next
            CRI = (CRI - CRITEST(8).Ri) * 0.125
            R9 = CRITEST(8).Ri
        Else
            CRI = -1000000
            R9 = -1000000
        End If
    End Sub
End Class
