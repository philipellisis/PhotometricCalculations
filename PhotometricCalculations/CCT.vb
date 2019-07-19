Public Class CCT
    Public Lumens As Double
    Public x As Double
    Public y As Double
    Public u As Double
    Public v As Double
    Public uPrime As Double
    Public vPrime As Double
    Public cct As Double
    ''' <summary>
    ''' calculates various CIE color coordinates
    ''' </summary>
    ''' <param name="spd">Ensure that the wavelength is from 380-780 in 1nm increments</param>
    Public Sub calculate(spd() As Double)
        'calculate variables x,y,u,v,uPrime,vPrime
        Dim bigX, bigY, bigZ As Double
        Dim cnt As Integer


        For cnt = 0 To 400
            bigX += CIE2DegreeFunctions.xbar(cnt) * spd(cnt)
            bigY += CIE2DegreeFunctions.ybar(cnt) * spd(cnt)
            bigZ += CIE2DegreeFunctions.zbar(cnt) * spd(cnt)
        Next cnt

        Try
            Lumens = bigY * 683
            x = bigX / (bigX + bigY + bigZ)
            y = bigY / (bigX + bigY + bigZ)
            u = 4 * x / (-2 * x + 12 * y + 3)
            v = 6 * y / (-2 * x + 12 * y + 3)
            uPrime = u
            vPrime = v * 1.5
            cct = ParseCCT(x, y)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Shared Function ParseCCT(ByVal x As Double, ByVal y As Double) As Double
        Dim closestuv As Integer
        Dim closestduv As Double = 100
        Dim a As Double
        Dim b As Double
        Dim c As Double
        Dim bigA As Double
        Dim bigB As Double
        Dim bigC As Double
        Dim tm As Double
        Dim tmPlus As Double
        Dim tmMinus As Double
        Dim dm As Double
        Dim dmPlus As Double
        Dim dmMinus As Double
        Dim u As Double
        Dim v As Double
        Dim CCT As Double = 0



        closestduv = 100
        closestuv = 0
        u = 4 * x / (-2 * x + 12 * y + 3)
        v = 6 * y / (-2 * x + 12 * y + 3)
        For k As Integer = 0 To 302 Step 1

            If closestduv > (u - CIE2DegreeFunctions.bbu(k)) ^ 2 + (v - CIE2DegreeFunctions.bbv(k)) ^ 2 Then
                closestduv = (((u - CIE2DegreeFunctions.bbu(k)) ^ 2 + (v - CIE2DegreeFunctions.bbv(k)) ^ 2))
                closestuv = k
            End If
        Next
        If closestuv > 0 And closestuv < 302 Then
            tm = CIE2DegreeFunctions.cct(closestuv)
            tmMinus = CIE2DegreeFunctions.cct(closestuv - 1)
            tmPlus = CIE2DegreeFunctions.cct(closestuv + 1)
            dm = closestduv
            dmPlus = ((u - CIE2DegreeFunctions.bbu(closestuv + 1)) ^ 2 + (v - CIE2DegreeFunctions.bbv(closestuv + 1)) ^ 2)
            dmMinus = ((u - CIE2DegreeFunctions.bbu(closestuv - 1)) ^ 2 + (v - CIE2DegreeFunctions.bbv(closestuv - 1)) ^ 2)

            a = dmMinus / (tmMinus - tm) / (tmMinus - tmPlus)
            b = dm / (tm - tmMinus) / (tm - tmPlus)
            c = dmPlus / (tmPlus - tmMinus) / (tmPlus - tm)
            bigA = a + b + c
            bigB = -1 * (a * (tmPlus + tm) + b * (tmMinus + tmPlus) + c * (tm + tmMinus))
            bigC = a * tm * tmPlus + b * tmMinus * tmPlus + c * tm * tmMinus
            CCT = -1 * bigB / (2 * bigA)
        Else
            CCT = 0
        End If
        Return CCT
    End Function
End Class
